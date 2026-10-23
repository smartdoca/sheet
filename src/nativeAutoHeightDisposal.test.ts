import {expect,it,vi} from 'vitest'
import {cancelAutoHeightOnDispose} from './nativeAutoHeightDisposal'

it('cancels the native idle callback and releases retained skeleton tasks before render disposal',()=>{
  const idle=new Map([[7,()=>{throw Error('Disposed sheet skeleton accessed')} ]])
  vi.stubGlobal('cancelIdleCallback',(id:number)=>idle.delete(id))
  try{
    const dispose=vi.fn(),service={dispose,_calculateLoopId:7 as number|null,_autoHeightTasks:[{skeleton:{}}]}
    cancelAutoHeightOnDispose(service);service.dispose()
    for(const task of idle.values())task()
    expect(service._calculateLoopId).toBeNull();expect(service._autoHeightTasks).toEqual([]);expect(dispose).toHaveBeenCalledOnce()
  }finally{vi.unstubAllGlobals()}
})
