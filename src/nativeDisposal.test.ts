import {expect,it,vi} from 'vitest'
import {BehaviorSubject,debounceTime} from 'rxjs'
import {clearHoverBeforeDispose} from './nativeDisposal'

it('invalidates a pending hover before workbook removal and completion flush, exactly once',()=>{
  const hover=new BehaviorSubject<{unitId:string}|null>(null)
  const workbook={getSheetBySheetId:vi.fn()},units=new Map([['history',workbook]])
  const seen:unknown[]=[]
  // This is the native alert's debounce and lookup ordering. Completion flushes
  // synchronously even when the 100ms timer has not elapsed.
  hover.pipe(debounceTime(100)).subscribe(position=>{
    seen.push(position)
    if(position)units.get(position.unitId)!.getSheetBySheetId()
  })
  const nativeDispose=vi.fn(()=>{units.clear();hover.complete()})
  const runtime={dispose:nativeDispose}
  clearHoverBeforeDispose(runtime,{triggerScroll:()=>hover.next(null)})
  hover.next({unitId:'history'})
  runtime.dispose();runtime.dispose()
  expect(seen).toEqual([null]);expect(nativeDispose).toHaveBeenCalledOnce()
  expect(workbook.getSheetBySheetId).not.toHaveBeenCalled()
})
