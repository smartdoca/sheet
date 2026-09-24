import { afterEach, describe, expect, it, vi } from 'vitest'
import * as Y from 'yjs'
import { createExlsxBaseline, restoreExlsxDocument, createExlsxCollaborationSession, type ExlsxRecoveryBundle, type ExlsxCollaborationSession, type ExlsxLocalTransaction } from './session'
import type { CollaborationContext, CollaborationMutation, WorkbookSnapshot } from './types'
import { inlineBody } from './inlineModel'

const snapshot = { id: 'book', name: 'test', styles: {}, sheetOrder: ['sheet'], sheets: { sheet: { id: 'sheet', name: 'Sheet', rowCount: 220, columnCount: 26, cellData: { 0: { 0: { v: 'baseline' } } } } } } as unknown as WorkbookSnapshot
const disposers: Array<() => void> = []
afterEach(() => { disposers.splice(0).forEach(dispose => dispose()); vi.useRealTimers() })
async function replica(bundle: ExlsxRecoveryBundle, sessionId: string) {
  const doc = await restoreExlsxDocument(bundle)
  const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId })
  const model = structuredClone(bundle.baseline.snapshot)
  const callbacks = new Set<(mutation: CollaborationMutation) => void>()
  const transactions: ExlsxLocalTransaction[] = []
  session.onLocalTransaction(event => transactions.push(event))
  const context = {
    workbookId: 'book', initialSnapshot: bundle.baseline.snapshot,
    getSnapshot: () => model,
    onLocalMutation(callback) { callbacks.add(callback); return () => { callbacks.delete(callback) } },
    async applyRemoteMutation(mutation) { apply(mutation); return true },
  } as CollaborationContext
  function apply(mutation: CollaborationMutation) {
    if(mutation.id==='sheet.mutation.set-frozen'){model.sheets.sheet.freeze=mutation.params as any;return}
    if(mutation.id!=='sheet.mutation.set-range-values')return
    const values = mutation.params?.cellValue as Record<string, Record<string, object>>
    for (const [r, row] of Object.entries(values)) for (const [c, cell] of Object.entries(row)) {
      const matrix = model.sheets.sheet.cellData ??= {}
      ;(matrix[Number(r)] ??= {})[Number(c)] = { ...matrix[Number(r)]?.[Number(c)], ...cell }
    }
  }
  const edit = (cellValue: Record<number, Record<number, unknown>>) => {
    const mutation = { id: 'sheet.mutation.set-range-values', params: { unitId: 'book', subUnitId: 'sheet', cellValue } }
    session.validateLocalMutation?.(mutation); apply(mutation); callbacks.forEach(callback => callback(mutation))
  }
  await session.connect(context); await session.ready
  disposers.push(() => { session.dispose(); doc.destroy() })
  return { doc, session, model, edit, transactions, context, callbacks }
}
async function deliver(from: { transactions: ExlsxLocalTransaction[] }, to: { session: ExlsxCollaborationSession }) {
  for (const transaction of from.transactions) await to.session.applyUpdate(transaction)
}
const value = (r: Awaited<ReturnType<typeof replica>>, column = 0) => r.model.sheets.sheet.cellData?.[0]?.[column]?.v

describe('host-managed exlsx session', () => {
  it('isolates CRDT and baseline content from mutable engine input and remote/undo projection',async()=>{
    const bundle=await createExlsxBaseline(snapshot,'engine-isolation'),a=await replica(bundle,'isolate-a'),b=await replica(bundle,'isolate-b')
    const p={id:'cell-doc',body:{dataStream:'文字\r\n'},documentStyle:{pageSize:{}}}
    a.edit({0:{0:{p,v:null}}})
    p.body.dataStream='外部修改不能污染寄存器\r\n'
    const apply=b.context.applyRemoteMutation
    b.context.applyRemoteMutation=async mutation=>{
      const cell=(mutation.params?.cellValue as any)?.[0]?.[0]
      if(cell?.p){cell.p.documentStyle.pageSize.width=Infinity;cell.p.body.dataStream='模拟引擎派生值';cell.customRender=()=>{}}
      return apply(mutation)
    }
    await deliver(a,b)
    const restored=await replica(b.session.checkpoint(1),'isolate-reload')
    expect(restored.model.sheets.sheet.cellData![0][0].p?.body?.dataStream).toBe('文字\r\n')
    const originalApply=a.context.applyRemoteMutation
    a.context.applyRemoteMutation=async mutation=>{
      const cell=(mutation.params?.cellValue as any)?.[0]?.[0]
      if(cell?.p)cell.p.documentStyle.pageSize.width=Infinity
      return originalApply(mutation)
    }
    await a.session.undo();await a.session.redo()
    expect((await replica(a.session.checkpoint(3),'isolate-undo-reload')).model.sheets.sheet.cellData![0][0].p?.body?.dataStream).toBe('文字\r\n')
    expect(bundle.baseline.snapshot).toEqual(snapshot)
  })
  it('preserves mixed document identity and safe links through default-schema replicas and checkpoint',async()=>{
    const bundle=await createExlsxBaseline(snapshot,'mixed-reference-epoch'),a=await replica(bundle,'reference-a'),b=await replica(bundle,'reference-b')
    const document=inlineBody({kind:'atomic',node:{type:'document',refId:'doc-123',label:'📄 需求文档'}},'doc-range')
    const link=inlineBody({kind:'link',text:'外部说明',href:'https://example.test/help'},'link-range')
    const prefix='请查看 ',between='，参考 ',suffix=' 后反馈'
    const docOffset=prefix.length,linkOffset=docOffset+document.dataStream.length+between.length
    const p={id:'mixed-cell',body:{dataStream:prefix+document.dataStream+between+link.dataStream+suffix+'\r\n',customRanges:[...document.customRanges!.map(r=>({...r,startIndex:r.startIndex+docOffset,endIndex:r.endIndex+docOffset})),...link.customRanges!.map(r=>({...r,startIndex:r.startIndex+linkOffset,endIndex:r.endIndex+linkOffset}))],textRuns:[{st:0,ed:prefix.length,ts:{bl:1}}]}}
    a.edit({0:{0:{p,v:null}}});await deliver(a,b)
    expect(b.transactions).toHaveLength(0)
    expect((await replica(b.session.checkpoint(1),'reference-reload')).model.sheets.sheet.cellData![0][0].p).toEqual(p)
    await a.session.undo();await a.session.redo();expect(a.model.sheets.sheet.cellData![0][0].p).toEqual(p)
    const unsafe=structuredClone(p);unsafe.body.customRanges[1].properties!.url='data:text/html,bad'
    expect(()=>a.edit({0:{0:{p:unsafe}}})).toThrow('UNSAFE_INLINE_LINK')
  })
  it('preserves native inline identity and partial styles through replicas, checkpoint and undo', async()=>{
    const bundle=await createExlsxBaseline(snapshot,'inline-epoch'),a=await replica(bundle,'inline-a'),b=await replica(bundle,'inline-b')
    const body=inlineBody({kind:'atomic',node:{type:'user',refId:'user-1',label:'@张三'}},'inline-range')
    body.dataStream+=' 后文\r\n';body.textRuns=[{st:4,ed:6,ts:{bl:1}}]
    const p={id:'native-cell',body}
    a.edit({0:{0:{p,v:null}}});await deliver(a,b)
    expect(b.model.sheets.sheet.cellData?.[0]?.[0]?.p).toEqual(p);expect(b.transactions).toHaveLength(0)
    const restored=await replica(a.session.checkpoint(1),'inline-restored')
    expect(restored.model.sheets.sheet.cellData?.[0]?.[0]?.p).toEqual(p)
    await a.session.undo();await a.session.redo()
    expect(a.model.sheets.sheet.cellData?.[0]?.[0]?.p).toEqual(p)
    const damaged=structuredClone(p);damaged.body.dataStream='@张X 后文\r\n'
    expect(()=>a.edit({0:{0:{p:damaged}}})).toThrow('INVALID_INLINE_RANGE')
    await deliver(a,b)
    const before=a.transactions.length
    b.edit({0:{0:{v:'later remote',p:null}}});await deliver(b,a);await a.session.undo()
    expect(value(a)).toBe('later remote');expect(a.transactions).toHaveLength(before)
  })
  it('normalizes native URL paste page dimensions without weakening JSON validation', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'paste-epoch')
    const a = await replica(bundle, 'paste-a'), b = await replica(bundle, 'paste-b')
    const p = { id: 'native-cell-document', body: { dataStream: 'https://example.test/#/r/doc\r\n' }, documentStyle: { pageSize: { width: Infinity, height: Infinity }, marginTop: 0 } }
    a.edit({ 0: { 0: { p } } })
    expect(p.documentStyle.pageSize.width).toBe(Infinity)
    expect(a.transactions).toHaveLength(1)
    await deliver(a, b)
    expect(b.transactions).toHaveLength(0)
    expect(b.model.sheets.sheet.cellData![0][0].p).toEqual({ ...p, documentStyle: { ...p.documentStyle, pageSize: {} } })
    const restored = await replica(a.session.checkpoint(1), 'paste-reload')
    expect(restored.model.sheets.sheet.cellData).toMatchObject(b.model.sheets.sheet.cellData!)
    a.edit({ 0: { 0: { p } } })
    expect(a.transactions).toHaveLength(1)
    for (const invalid of [NaN, -Infinity]) {
      expect(() => a.edit({ 0: { 0: { p: { ...p, documentStyle: { pageSize: { width: invalid } } } } } })).toThrow('JSON data')
    }
    expect(() => a.edit({ 0: { 0: { p: { ...p, unexpected: Infinity } } } })).toThrow('JSON data')
    expect(() => a.edit({ 0: { 0: { p: { ...p, unexpected: new Uint8Array([1]) } } } })).toThrow('JSON data')
    expect(a.transactions).toHaveLength(1)
    await a.session.undo(); await b.session.applyUpdate(a.transactions.at(-1)!)
    expect(value(b)).toBe('baseline')
    await a.session.redo(); await b.session.applyUpdate(a.transactions.at(-1)!)
    expect(b.model.sheets.sheet.cellData![0][0].p).toEqual(restored.model.sheets.sheet.cellData![0][0].p)
  })
  it('keeps incremental validation isolated after rejection, local edits, undo, and duplicate delivery', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'validator-a'); const b = await replica(bundle, 'validator-b')
    a.edit({0:{0:{v:'warm'}}}); await deliver(a,b)
    b.edit({0:{1:{v:'local mirrored'}}}); await b.session.undo(); await b.session.redo()
    const evil = await restoreExlsxDocument(b.session.checkpoint(3))
    const vector = Y.encodeStateVector(evil)
    evil.getMap('exlsx:cells').set('["sheet",0,0,"content"]', {v:'evil',f:null,p:{body:{customBlocks:[{blockId:'x'}]}},t:null,si:null})
    const malicious = {...a.transactions[0],update:Y.encodeStateAsUpdate(evil,vector)}
    await expect(b.session.applyUpdate(malicious)).rejects.toMatchObject({code:'INVALID_UPDATE'})
    evil.destroy()
    expect(value(b)).toBe('warm');expect(value(b,1)).toBe('local mirrored')
    const before = b.transactions.length
    a.edit({0:{0:{v:'after rejection'}}}); await b.session.applyUpdate(a.transactions.at(-1)!);await b.session.applyUpdate(a.transactions.at(-1)!)
    expect(value(b)).toBe('after rejection');expect(value(b,1)).toBe('local mirrored');expect(b.transactions).toHaveLength(before)
    expect((await replica(b.session.checkpoint(4),'restored-validator')).model.sheets.sheet.cellData).toMatchObject(b.model.sheets.sheet.cellData!)
  })
  it('roundtrips whole-cell business objects with one local transaction and no remote echo', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'object-epoch');
    const a = await replica(bundle, 'a'), b = await replica(bundle, 'b');
    const object = { version: 1, kind: 'mention', id: 'user-1', label: '用户' };
    a.edit({ 0: { 1: { v: '@用户', custom: { officeObject: object } } } });
    expect(a.transactions).toHaveLength(1); await deliver(a, b);
    expect(b.model.sheets.sheet.cellData![0][1]).toMatchObject({ v: '@用户', custom: { officeObject: object } });
    expect(b.transactions).toHaveLength(0);
    const restored = await replica(a.session.checkpoint(1), 'restored');
    expect(restored.model.sheets.sheet.cellData![0][1]).toMatchObject({ custom: { officeObject: object } });
    await a.session.undo(); expect(a.model.sheets.sheet.cellData![0][1]?.custom).toBeNull();
    await a.session.redo(); expect(a.model.sheets.sheet.cellData![0][1]?.custom).toEqual({ officeObject: object });
    a.edit({ 0: { 1: null } }); await deliver(a, b); expect(b.model.sheets.sheet.cellData![0][1]?.custom).toBeNull();
  });
  it('ignores runtime renderer caches in local edits without persisting or echoing them', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'a'); const b = await replica(bundle, 'b')
    const customRender = [{ drawWith() {} }]
    a.edit({ 0: { 0: { v: 'edited', customRender } } })
    expect(a.transactions).toHaveLength(1)
    await deliver(a, b)
    expect(value(b)).toBe('edited')
    expect(b.model.sheets.sheet.cellData![0][0]).not.toHaveProperty('customRender')
    expect(b.transactions).toHaveLength(0)
    a.edit({ 0: { 0: { customRender } } })
    expect(a.transactions).toHaveLength(1)
    expect(() => a.edit({ 0: { 0: { unexpected: true } } })).toThrow('Unsupported cell field')
    expect(a.session.state).toBe('ready')
  })
  it('rejects malformed custom-block identity nodes before a local content transaction', async () => {
    const r = await replica(await createExlsxBaseline(snapshot, 'epoch'), 'inline-probe')
    expect(() => r.edit({ 0: { 0: { p: { id: 'rich', body: { dataStream: '\b\r\n', customBlocks: [{ startIndex: 0, blockId: 'user-node' }] }, mentions: { 'user-node': { objectId: 'user-1', label: '用户' } } } } } })).toThrow('INVALID_INLINE_DRAWING_MEMBERSHIP')
    expect(r.transactions).toEqual([])
    expect(value(r)).toBe('baseline')
    r.edit({ 0: { 0: { v: '你好 @用户' } } })
    const restored = await replica(r.session.checkpoint(1), 'plain-text-reload')
    expect(value(restored)).toBe('你好 @用户')
    expect(restored.model.sheets.sheet.cellData?.[0]?.[0]?.p).toBeFalsy()
  })
  it('emits zero transactions during a 60-second idle interval and readonly toggles', async () => {
    const r = await replica(await createExlsxBaseline(snapshot, 'epoch'), 'same-user-tab-a')
    vi.useFakeTimers(); r.session.setReadOnly(true); r.session.setReadOnly(false)
    await vi.advanceTimersByTimeAsync(60_000)
    expect(r.transactions).toEqual([])
  })
  it('converges same-account distinct sessions with no remote echo, duplicate updates or pure-delete sync writes', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'user1-tab1'); const b = await replica(bundle, 'user1-tab2')
    a.edit({ 0: { 0: { v: 'A' } } }); b.edit({ 0: { 1: { v: 'B' } } })
    await deliver(a, b); await deliver(b, a); await deliver(a, b)
    expect([value(a), value(a, 1)]).toEqual(['A', 'B'])
    expect([value(b), value(b, 1)]).toEqual(['A', 'B'])
    expect([a.transactions.length, b.transactions.length]).toEqual([1, 1])
    a.edit({ 0: { 0: null } }); await deliver(a, b)
    for (let i = 0; i < 3; i++) await b.session.applyUpdate({ ...a.transactions[0], update: Y.encodeStateAsUpdate(a.doc, Y.encodeStateVector(b.doc)) })
    expect(b.transactions).toHaveLength(1); expect(value(b)).toBeNull()
  })
  it('resolves same-cell formula/value races atomically and deterministically', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'a'); const b = await replica(bundle, 'b')
    a.edit({ 0: { 0: { f: '=1+2', v: null, p: null } } })
    b.edit({ 0: { 0: { f: null, v: 'text', p: null } } })
    await deliver(a, b); await deliver(b, a)
    expect(a.model.sheets.sheet.cellData).toEqual(b.model.sheets.sheet.cellData)
    const cell = a.model.sheets.sheet.cellData![0][0]!
    expect(cell.f === '=1+2' ? cell.v === null : cell.v === 'text').toBe(true)
  })
  it('undo never overwrites a later remote winner and redo only restores own eligible changes', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'a'); const b = await replica(bundle, 'b')
    a.edit({ 0: { 0: { v: 'A' } } }); await deliver(a, b)
    b.edit({ 0: { 0: { v: 'B' } } }); await deliver(b, a)
    await a.session.undo(); await deliver(a, b)
    expect(value(a)).toBe('B'); expect(value(b)).toBe('B')
    a.edit({ 0: { 1: { v: 'own' } } }); await a.session.undo()
    expect(value(a, 1)).toBeNull()
    await a.session.redo(); expect(value(a, 1)).toBe('own')
    expect(a.transactions.at(-1)?.kind).toBe('redo')
  })
  it('one multi-cell mutation produces one transaction and one undo step', async () => {
    const r = await replica(await createExlsxBaseline(snapshot, 'epoch'), 'a')
    r.edit({ 0: { 0: { v: 'replace' }, 1: { v: 'replace' } } })
    expect(r.transactions).toHaveLength(1)
    await r.session.undo(); expect(value(r)).toBe('baseline'); expect(value(r, 1)).toBeNull()
  })
  it('restores checkpoints and merged updates without changing epoch, baseline or permanent anchors', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'a')
    const anchor = a.session.captureCellAnchor!({ sheetId: 'sheet', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 })!
    a.edit({ 0: { 0: { v: 'first' } } }); a.edit({ 0: { 0: { v: 'last' } } })
    const checkpoint = a.session.checkpoint(2)
    const b = await replica(checkpoint, 'b')
    const c = await replica({ ...bundle, update: Y.mergeUpdates([bundle.update, ...a.transactions.map(t => t.update)]), checkpointSeq: 2 }, 'c')
    expect(value(b)).toBe('last'); expect(value(c)).toBe('last')
    expect(b.session.captureCellAnchor!(b.session.resolveCellAnchor!(anchor)!)).toEqual(anchor)
    expect(checkpoint.baseline).toEqual(bundle.baseline)
    expect(b.transactions).toEqual([])
    expect(b.session.resolveCellAnchor!({ ...anchor, epochId: 'different' })).toBeNull()
  })
  it('rejects mismatched epoch/schema/baseline without overwriting pending local changes', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'a'); a.edit({ 0: { 0: { v: 'pending' } } })
    await expect(a.session.applyUpdate({ ...a.transactions[0], epochId: 'other' })).rejects.toMatchObject({ code: 'EPOCH_MISMATCH' })
    await expect(a.session.applyUpdate({ ...a.transactions[0], schemaVersion: 99 } as never)).rejects.toMatchObject({ code: 'SCHEMA_MISMATCH' })
    const wrong = structuredClone(bundle); wrong.baseline.snapshot.name = 'different'
    await expect(restoreExlsxDocument(wrong)).rejects.toMatchObject({ code: 'BASELINE_MISMATCH' })
    expect(value(a)).toBe('pending'); expect(a.transactions).toHaveLength(1)
  })
  it('preserves explicit schema 1 restrictions without silently upgrading old epochs', async () => {
    const r = await replica(await createExlsxBaseline(snapshot, 'epoch', {schemaVersion:1}), 'a')
    for (const id of ['insert-row', 'remove-rows', 'insert-col', 'remove-col', 'add-worksheet-merge', 'remove-worksheet-merge', 'sort-range', 'set-filter-range']) {
      const mutation = { id: `sheet.mutation.${id}` }
      expect(r.session.supportsMutation!(mutation.id)).toBe(false)
      expect(() => r.session.validateLocalMutation!(mutation)).toThrow('does not support')
    }
    expect(r.transactions).toEqual([])
  })
  it('readonly rejects writes/undo but continues projecting remote updates; disposal leaves host Doc alive', async () => {
    const bundle = await createExlsxBaseline(snapshot, 'epoch')
    const a = await replica(bundle, 'a'); const b = await replica(bundle, 'b')
    b.session.setReadOnly(true)
    expect(() => b.edit({ 0: { 0: { v: 'bad' } } })).toThrow('read only')
    await expect(b.session.undo()).rejects.toMatchObject({ code: 'READ_ONLY' })
    a.edit({ 0: { 0: { v: 'remote' } } }); await deliver(a, b)
    expect(value(b)).toBe('remote'); expect(b.transactions).toEqual([])
    b.session.setReadOnly(false); b.edit({ 0: { 1: { v: 'now allowed' } } })
    b.session.dispose(); expect(b.doc.isDestroyed).toBe(false); expect(b.callbacks.size).toBe(0)
  })
})
