import * as Y from 'yjs'

/** Sparse identity sequence. Baseline identities are virtual, never inserted into Yjs.
 * Positions belong to identities, not screen coordinates. Deletion claims are per
 * session so undoing my deletion cannot revive a row another session deleted.
 */
export const AXIS_POSITIONS = 'exlsx:axis-positions'
export const AXIS_DELETIONS = 'exlsx:axis-deletions'
/** Exclude this immutable catalog from UndoManager. Undone insertions remain
 * identifiable so concurrent edits can be retained without appearing elsewhere. */
export const AXIS_IDENTITIES = 'exlsx:axis-identities'
export type Axis = 'row' | 'column'
export type AxisId = string
export type AxisPosition = readonly number[]
export interface AxisAddress { sheetId: string; axis: Axis; id: AxisId }
const MAX_DIGIT = 0x40000000
export const AXIS_LIMITS = Object.freeze({ rows: 1_048_576, columns: 16_384, batch: 10_000, positionDepth: 128 })
export const axisKey = ({sheetId, axis, id}: AxisAddress) => JSON.stringify([sheetId, axis, id])
export function axisAddress(key: string): AxisAddress {
  const a = JSON.parse(key)
  if (!Array.isArray(a) || a.length !== 3 || typeof a[0] !== 'string' || !a[0] || !['row','column'].includes(a[1]) || !validAxisId(a[2]) || JSON.stringify(a) !== key) throw new Error('INVALID_AXIS_KEY')
  return {sheetId:a[0], axis:a[1], id:a[2]}
}
export function validAxisId(id: unknown): id is AxisId {
  return typeof id === 'string' && (/^[br]:(0|[1-9]\d*)$/.test(id) || /^i:[a-f0-9-]{36}$/.test(id))
}
export function comparePosition(a: AxisPosition, b: AxisPosition): number {
  for (let i=0; i<Math.min(a.length,b.length); i++) if (a[i] !== b[i]) return a[i]-b[i]
  return a.length-b.length
}
export function validatePosition(value: unknown): asserts value is AxisPosition {
  if (!Array.isArray(value) || !value.length || value.length>AXIS_LIMITS.positionDepth || value.some(n=>!Number.isSafeInteger(n)||n<0||n>=MAX_DIGIT) || value.at(-1)===0) throw new Error('INVALID_AXIS_POSITION')
}
/** UUID-derived suffix prevents concurrent positions colliding; no clock or user ID. */
export function positionBetween(left: AxisPosition, right: AxisPosition, uuid: string): AxisPosition {
  if (comparePosition(left,right)>=0) throw new Error('INVALID_AXIS_NEIGHBORS')
  const out:number[]=[]
  let bounded=true
  for(let i=0;i<AXIS_LIMITS.positionDepth-10;i++) {
    const l=left[i]??0, r=bounded?(right[i]??MAX_DIGIT):MAX_DIGIT
    if(r-l>1) {
      out.push(l+Math.floor((r-l)/2))
      const hex=uuid.replaceAll('-','')
      if(!/^[a-f0-9]{32}$/.test(hex))throw new Error('INVALID_AXIS_UUID')
      for(let j=0;j<hex.length;j+=4)out.push(parseInt(hex.slice(j,j+4),16)+1)
      validatePosition(out)
      if(comparePosition(left,out)>=0||comparePosition(out,right)>=0)throw new Error('AXIS_POSITION_EXHAUSTED')
      return out
    }
    out.push(l)
    if(l<r)bounded=false
  }
  throw new Error('AXIS_POSITION_EXHAUSTED')
}
/** Balanced allocation avoids exhausting fractional precision on a 10k-row paste
 * or repeated sorts. Sequential midpoint allocation grows linearly in count. */
function allocatePositions(left:AxisPosition,right:AxisPosition,uuids:readonly string[]):AxisPosition[]{
  const result:AxisPosition[]=new Array(uuids.length)
  const fill=(start:number,end:number,l:AxisPosition,r:AxisPosition)=>{if(start>=end)return;const mid=Math.floor((start+end)/2),p=positionBetween(l,r,uuids[mid]);result[mid]=p;fill(start,mid,l,p);fill(mid+1,end,p,r)}
  fill(0,uuids.length,left,right);return result
}

export interface StableAxis {
  readonly length: number
  idAt(index:number): AxisId | null
  indexOf(id:AxisId): number
  contains(id:AxisId): boolean
  ids(start:number,end:number): AxisId[]
  position(id:AxisId): AxisPosition | null
  insert(index:number,count:number): AxisId[]
  remove(start:number,count:number,sessionId:string): AxisId[]
  /** Reorder complete records. Each identity's position is an independent LWW register. */
  reorder(ids:readonly AxisId[],start:number):void
  dispose():void
}

/** Caller must wrap writes in its normal session-origin transaction and enforce ACL.
 * Reading a pristine million-row axis is O(1) and allocates no identity array.
 * A structural revision builds one cached index, reused by every cell projection.
 */
export function createStableAxis(doc:Y.Doc,sheetId:string,axis:Axis,baselineCount:number):StableAxis {
  const max=axis==='row'?AXIS_LIMITS.rows:AXIS_LIMITS.columns
  if(!Number.isSafeInteger(baselineCount)||baselineCount<1||baselineCount>max)throw new Error('INVALID_AXIS_BASELINE')
  const positions=doc.getMap<AxisPosition>(AXIS_POSITIONS),deletions=doc.getMap<boolean>(AXIS_DELETIONS),identities=doc.getMap<boolean>(AXIS_IDENTITIES)
  const prefix=JSON.stringify([sheetId,axis]).slice(0,-1)+','
  type Chunk={offset:number;length:number;start?:number;id?:string}
  let dirty=true,chunks:Chunk[]=[],baseChunks:Chunk[]=[],length=baselineCount,inverse=new Map<AxisId,number>()
  let overrides=new Map<AxisId,AxisPosition>(),deleted=new Set<AxisId>()
  let recoveryGeneration=0
  const baseIndex=(id:string)=>/^b:(0|[1-9]\d*)$/.test(id)?Number(id.slice(2)):-1
  const known=(id:string)=>{const i=baseIndex(id);return i>=0?i<baselineCount:/^r:(0|[1-9]\d*)$/.test(id)?Number(id.slice(2))<=recoveryGeneration:identities.get(axisKey({sheetId,axis,id}))===true}
  const pos=(id:string):AxisPosition|null=>overrides.get(id)??(id.startsWith('r:')?[1]:baseIndex(id)>=0&&baseIndex(id)<baselineCount?[2+baseIndex(id)*2]:null)
  function rebuild(){
    if(!dirty)return
    overrides=new Map();deleted=new Set()
    recoveryGeneration=0
    for(const [key,value] of deletions){const a=JSON.parse(key);if(a[0]===sheetId&&a[1]===axis&&value===true)deleted.add(a[2])}
    for(const id of deleted)if(/^r:(0|[1-9]\d*)$/.test(id)){const n=Number(id.slice(2));if(!Number.isSafeInteger(n)||n>1_000_000)throw new Error('INVALID_RECOVERY_AXIS_ID');recoveryGeneration=Math.max(recoveryGeneration,n+1)}
    for(const [key,value] of positions)if(key.startsWith(prefix)){const a=axisAddress(key);validatePosition(value);if(!known(a.id))throw new Error('UNKNOWN_AXIS_ID');overrides.set(a.id,value)}
    for(const id of deleted)if(!known(id))throw new Error('UNKNOWN_DELETED_AXIS_ID')
    // Keep unchanged baseline runs as intervals. A single edit in a million-row
    // sheet must not allocate a million strings and inverse-map entries.
    const changed=[...overrides].filter(([id])=>!deleted.has(id)).sort((a,b)=>comparePosition(a[1],b[1])||a[0].localeCompare(b[0]))
    const excluded=[...new Set([...overrides.keys(),...deleted].map(baseIndex).filter(i=>i>=0))].sort((a,b)=>a-b)
    chunks=[];baseChunks=[];inverse=new Map();length=0
    let cursor=0,skip=0
    const append=(end:number)=>{
      while(skip<excluded.length&&excluded[skip]<cursor)skip++
      while(cursor<end){
        if(excluded[skip]===cursor){cursor++;skip++;continue}
        const next=Math.min(end,excluded[skip]??end),chunk={offset:length,length:next-cursor,start:cursor}
        chunks.push(chunk);baseChunks.push(chunk);length+=chunk.length;cursor=next
      }
    }
    for(const [id,position] of changed){
      let lo=cursor,hi=baselineCount
      while(lo<hi){const mid=Math.floor((lo+hi)/2),cmp=comparePosition([2+mid*2],position);if(cmp<0||cmp===0&&`b:${mid}`.localeCompare(id)<0)lo=mid+1;else hi=mid}
      append(lo);inverse.set(id,length);chunks.push({offset:length++,length:1,id})
    }
    append(baselineCount)
    // Concurrent disjoint deletions may collectively remove every record. Render
    // one deterministic NEW blank identity instead of reviving a deleted record
    // (which would resurrect comments) or rejecting an otherwise valid CRDT merge.
    if(!length){const id=`r:${recoveryGeneration}`;chunks.push({offset:0,length:1,id});inverse.set(id,0);length=1}
    if(length>max)throw new Error('AXIS_SIZE_LIMIT')
    dirty=false
  }
  const changed=(event:{keysChanged:Set<string>})=>{for(const key of event.keysChanged)if(key.startsWith(prefix)){dirty=true;break}}
  positions.observe(changed);deletions.observe(changed)
  const api:StableAxis={
    get length(){rebuild();return length},
    idAt(index){
      rebuild();if(!Number.isSafeInteger(index)||index<0||index>=length)return null
      let lo=0,hi=chunks.length
      while(lo<hi){const mid=Math.floor((lo+hi)/2);if(chunks[mid].offset+chunks[mid].length<=index)lo=mid+1;else hi=mid}
      const chunk=chunks[lo];return chunk.id??`b:${chunk.start!+index-chunk.offset}`
    },
    indexOf(id){
      rebuild();if(inverse.has(id))return inverse.get(id)!
      const index=baseIndex(id);if(index<0)return -1
      let lo=0,hi=baseChunks.length
      while(lo<hi){const mid=Math.floor((lo+hi)/2);if(baseChunks[mid].start!+baseChunks[mid].length<=index)lo=mid+1;else hi=mid}
      const chunk=baseChunks[lo];return chunk&&index>=chunk.start!?chunk.offset+index-chunk.start!:-1
    },
    contains(id){rebuild();return known(id)},
    ids(start,end){if(!Number.isSafeInteger(start)||!Number.isSafeInteger(end)||start<0||end<start||end>=api.length||end-start>=AXIS_LIMITS.batch)throw new Error('INVALID_AXIS_RANGE');return Array.from({length:end-start+1},(_,i)=>api.idAt(start+i)!)},
    position(id){rebuild();return pos(id)},
    insert(index,count){
      if(!Number.isSafeInteger(index)||index<0||index>api.length||!Number.isSafeInteger(count)||count<1||count>AXIS_LIMITS.batch||api.length+count>max)throw new Error('INVALID_AXIS_INSERT')
      const left=index?api.position(api.idAt(index-1)!)!:[0]
      const right=index<api.length?api.position(api.idAt(index)!)!:[MAX_DIGIT-1]
      const uuids=Array.from({length:count},()=>crypto.randomUUID()),slots=allocatePositions(left,right,uuids)
      const writes=uuids.map((uuid,i)=>({id:`i:${uuid}`,position:slots[i]}))
      const recovery=api.length===1?api.idAt(0):null
      if(recovery?.startsWith('r:')&&!overrides.has(recovery))positions.set(axisKey({sheetId,axis,id:recovery}),[1])
      for(const {id,position} of writes){const key=axisKey({sheetId,axis,id});identities.set(key,true);positions.set(key,position)}
      dirty=true;return writes.map(w=>w.id)
    },
    remove(start,count,sessionId){
      if(!sessionId||sessionId.length>256||!Number.isSafeInteger(count)||count<1||count>=api.length)throw new Error('INVALID_AXIS_DELETE')
      const ids=api.ids(start,start+count-1)
      for(const id of ids)deletions.set(JSON.stringify([sheetId,axis,id,sessionId]),true)
      dirty=true;return ids
    },
    reorder(ids,start){
      const current=api.ids(start,start+ids.length-1)
      const members=new Set(current)
      if(new Set(ids).size!==ids.length||ids.some(id=>!members.has(id)))throw new Error('INVALID_AXIS_PERMUTATION')
      const left=start?api.position(api.idAt(start-1)!)!:[0],right=start+ids.length<api.length?api.position(api.idAt(start+ids.length)!)!:[MAX_DIGIT-1]
      const slots=allocatePositions(left,right,ids.map(()=>crypto.randomUUID()))
      ids.forEach((id,i)=>positions.set(axisKey({sheetId,axis,id}),[...slots[i]]));dirty=true
    },
    dispose(){positions.unobserve(changed);deletions.unobserve(changed)},
  }
  return api
}
