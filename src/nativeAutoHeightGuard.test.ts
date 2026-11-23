import {expect,it,vi} from 'vitest'
import {UniverInstanceType,type Injector} from '@univerjs/core'
import {IRenderManagerService} from '@univerjs/engine-render'
import {Subject} from 'rxjs'
import {attachNativeAutoHeightGuard} from './nativeAutoHeightGuard'

function harness(renderers:unknown[]) {
  const created=new Subject()
  const manager={created$:created,getAllRenderersOfType:()=>renderers}
  const injector={get:(token:unknown)=>token===IRenderManagerService?manager:undefined} as Injector
  return {injector,created}
}

it('guards the skeleton measurement so one degenerate cell cannot cancel command chains',()=>{
  class Skeleton { calculateAutoHeightForCell(row:number,_col:number){ if(row===36) throw new TypeError("Cannot destructure property 'marginTop' of 'e.pages[(e.pages.length - 1)]' as it is undefined."); return 24 } }
  const skeleton=new Skeleton()
  const renderer={with:(token:unknown)=>({getCurrentSkeleton:()=>skeleton})}
  const {injector}=harness([renderer])
  const warn=vi.spyOn(console,'warn').mockImplementation(()=>{})
  try{
    expect(()=>skeleton.calculateAutoHeightForCell(36,0)).toThrow('marginTop')
    const adapter=attachNativeAutoHeightGuard(injector)
    expect(skeleton.calculateAutoHeightForCell(0,0)).toBe(24)
    expect(skeleton.calculateAutoHeightForCell(36,0)).toBeUndefined()
    expect(skeleton.calculateAutoHeightForCell(36,1)).toBeUndefined()
    expect(warn).toHaveBeenCalledTimes(1)
    adapter.dispose()
  }finally{warn.mockRestore()}
})

it('attaches to renderers registered later and never double-wraps the same prototype',async()=>{
  class Skeleton { calculateAutoHeightForCell(_row:number,_col:number){ return 24 } }
  const skeletons=[new Skeleton(),new Skeleton()]
  const renderers=[{with:(token:unknown)=>({getCurrentSkeleton:()=>skeletons[0]})}]
  const {injector,created}=harness(renderers)
  const adapter=attachNativeAutoHeightGuard(injector)
  renderers.push({with:(token:unknown)=>({getCurrentSkeleton:()=>skeletons[1]})})
  created.next(null)
  await Promise.resolve()
  const guarded=skeletons[0].calculateAutoHeightForCell
  created.next(null)
  await Promise.resolve()
  adapter.refresh()
  expect(skeletons[0].calculateAutoHeightForCell).toBe(guarded)
  expect(skeletons[1].calculateAutoHeightForCell).toBe(guarded)
  expect(skeletons[0].calculateAutoHeightForCell(1,0)).toBe(24)
  adapter.dispose()
})

it('keeps retrying until the skeleton exists, then stops the retry timer',()=>{
  vi.useFakeTimers()
  try{
    class Skeleton { calculateAutoHeightForCell(row:number,_col:number){ if(row===36) throw new TypeError('boom'); return 24 } }
    const holder:{current:Skeleton|null}={current:null}
    const renderers=[{with:(token:unknown)=>({getCurrentSkeleton:()=>holder.current})}]
    const {injector}=harness(renderers)
    const intervalSpy=vi.spyOn(globalThis,'clearInterval')
    const adapter=attachNativeAutoHeightGuard(injector)
    holder.current=new Skeleton()
    expect(()=>holder.current!.calculateAutoHeightForCell(36,0)).toThrow('boom')
    vi.advanceTimersByTime(600)
    expect(()=>holder.current!.calculateAutoHeightForCell(36,0)).not.toThrow()
    expect(intervalSpy).toHaveBeenCalled()
    adapter.dispose()
  }finally{vi.useRealTimers()}
})
