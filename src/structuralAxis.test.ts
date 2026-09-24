import {describe,it,expect} from 'vitest'
import * as Y from 'yjs'
import {AXIS_POSITIONS,AXIS_DELETIONS,createStableAxis,comparePosition,positionBetween,axisAddress} from './structuralAxis'

describe('sparse structural identities',()=>{
  it('interval indexes match a dense oracle across mixed insert/delete/reorder operations',()=>{
    const doc=new Y.Doc(),a=createStableAxis(doc,'s','row',30)
    let seed=31
    const random=(n:number)=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed%n}
    for(let step=0;step<160;step++){
      const choice=random(3)
      if(choice===0)a.insert(random(a.length+1),1+random(3))
      else if(choice===1&&a.length>3)a.remove(random(a.length),1,'session')
      else if(a.length>3){const start=random(a.length-2);a.reorder(a.ids(start,start+2).reverse(),start)}
      const overrides=new Map([...doc.getMap<readonly number[]>(AXIS_POSITIONS)].map(([k,v])=>[axisAddress(k).id,v]))
      const deleted=new Set([...doc.getMap(AXIS_DELETIONS)].map(([k])=>JSON.parse(k)[2]))
      const oracle=[...new Set([...Array.from({length:30},(_,i)=>`b:${i}`),...overrides.keys()])].filter(id=>!deleted.has(id)).sort((x,y)=>comparePosition(overrides.get(x)??[2+Number(x.slice(2))*2],overrides.get(y)??[2+Number(y.slice(2))*2])||x.localeCompare(y))
      expect(a.ids(0,a.length-1)).toEqual(oracle)
      for(const [i,id] of oracle.entries())expect(a.indexOf(id)).toBe(i)
    }
    a.dispose();doc.destroy()
  })
  it('inserts 10,000 identities and repeatedly sorts without exhausting position depth',()=>{
    const doc=new Y.Doc(),a=createStableAxis(doc,'s','row',2)
    const inserted=a.insert(1,10_000)
    expect(a.length).toBe(10_002);expect(new Set(inserted).size).toBe(10_000)
    const first=inserted.slice(0,100)
    for(let i=0;i<150;i++){first.reverse();a.reorder(first,1)}
    expect(a.ids(1,100)).toEqual(first)
    expect(a.insert(50,1)).toHaveLength(1)
    a.dispose();doc.destroy()
  })
  it('concurrent complete deletion creates a new empty identity, never resurrects old comments',()=>{
    const ad=new Y.Doc(),bd=new Y.Doc(),a=createStableAxis(ad,'s','row',2),b=createStableAxis(bd,'s','row',2)
    a.remove(0,1,'a');b.remove(1,1,'b');const au=Y.encodeStateAsUpdate(ad),bu=Y.encodeStateAsUpdate(bd)
    Y.applyUpdate(ad,bu);Y.applyUpdate(bd,au)
    expect(a.ids(0,0)).toEqual(['r:0']);expect(b.ids(0,0)).toEqual(['r:0']);expect(a.indexOf('b:0')).toBe(-1);expect(a.indexOf('b:1')).toBe(-1)
    a.insert(1,1);expect(a.idAt(0)).toBe('r:0');expect(a.length).toBe(2)
    a.remove(0,1,'a');expect(a.indexOf('r:0')).toBe(-1)
    a.dispose();b.dispose();ad.destroy();bd.destroy()
  })
  it('keeps a million baseline rows virtual',()=>{
    const doc=new Y.Doc(),a=createStableAxis(doc,'s','row',1_000_000)
    expect(a.idAt(999999)).toBe('b:999999');expect(a.indexOf('b:999999')).toBe(999999)
    expect(doc.getMap(AXIS_POSITIONS).size).toBe(0);expect(Y.encodeStateAsUpdate(doc).length).toBe(2)
    a.dispose();doc.destroy()
  })
  it('concurrent insertion, deletion and repeated delivery retain unique stable identities',()=>{
    const ad=new Y.Doc(),bd=new Y.Doc(),a=createStableAxis(ad,'s','row',6),b=createStableAxis(bd,'s','row',6)
    const ai=a.insert(2,2),bi=b.insert(2,2)
    a.remove(a.indexOf('b:3'),1,'a');b.remove(b.indexOf('b:3'),1,'b')
    const au=Y.encodeStateAsUpdate(ad),bu=Y.encodeStateAsUpdate(bd)
    Y.applyUpdate(ad,bu);Y.applyUpdate(bd,au);Y.applyUpdate(bd,au)
    expect(a.ids(0,a.length-1)).toEqual(b.ids(0,b.length-1));expect(a.length).toBe(9)
    expect(a.indexOf('b:3')).toBe(-1);expect(a.contains('b:3')).toBe(true)
    for(const id of [...ai,...bi])expect(a.indexOf(id)).toBeGreaterThan(0)
    a.dispose();b.dispose();ad.destroy();bd.destroy()
  })
  it('local undo cannot revive another session deletion, and checkpoint preserves tombstones',()=>{
    const ad=new Y.Doc(),bd=new Y.Doc(),a=createStableAxis(ad,'s','column',6),b=createStableAxis(bd,'s','column',6),origin={}
    const undo=new Y.UndoManager([ad.getMap(AXIS_POSITIONS),ad.getMap(AXIS_DELETIONS)],{trackedOrigins:new Set([origin])})
    ad.transact(()=>a.remove(2,1,'a'),origin);bd.transact(()=>b.remove(2,1,'b'))
    Y.applyUpdate(ad,Y.encodeStateAsUpdate(bd));undo.undo()
    expect(a.indexOf('b:2')).toBe(-1)
    const restored=new Y.Doc();Y.applyUpdate(restored,Y.encodeStateAsUpdate(ad));const c=createStableAxis(restored,'s','column',6)
    expect(c.ids(0,c.length-1)).toEqual(a.ids(0,a.length-1))
    undo.destroy();a.dispose();b.dispose();c.dispose();ad.destroy();bd.destroy();restored.destroy()
  })
  it('sort moves identities rather than contents and insertion works between concurrent positions',()=>{
    const doc=new Y.Doc(),a=createStableAxis(doc,'s','row',6)
    a.reorder(['b:3','b:2','b:1'],1);expect(a.ids(0,5)).toEqual(['b:0','b:3','b:2','b:1','b:4','b:5'])
    const [id]=a.insert(2,1);expect(a.idAt(2)).toBe(id);expect(a.idAt(3)).toBe('b:2')
    const l=positionBetween([2],[4],crypto.randomUUID()),r=positionBetween([2],[4],crypto.randomUUID()),[lo,hi]=[l,r].sort(comparePosition)
    const mid=positionBetween(lo,hi,crypto.randomUUID());expect(comparePosition(lo,mid)).toBeLessThan(0);expect(comparePosition(mid,hi)).toBeLessThan(0)
    a.dispose();doc.destroy()
  })
})
