import {expect,it,vi} from 'vitest'
import {UniverInstanceType,type Injector} from '@univerjs/core'
import {Subject} from 'rxjs'
import {attachNativeFocusOwner,holdNativeTextFocus} from './nativeFocus'
import type {SpreadsheetNativeText} from './inlineTypes'

it('popup ownership blocks all native focus entry points, allows explicit handoff, and restores instance methods',()=>{
  const created$=new Subject<any>(),disposed$=new Subject<string>()
  const service=()=>({focus:vi.fn(),activate(){this.focus()},sync(){this.activate()},addDocRanges(){this.sync()}})
  const a=service(),b=service(),original=a.focus,nativeB=b.focus,otherRuntime=service()
  const render=(id:string,s:typeof a,type=UniverInstanceType.UNIVER_DOC)=>({unitId:id,type,with:()=>s})
  const first=render('draft',a),sheet=render('sheet',otherRuntime,UniverInstanceType.UNIVER_SHEET)
  const manager={created$,disposed$,getRenderAll:()=>new Map([['draft',first],['sheet',sheet]])}
  const api={focus:()=>true} as SpreadsheetNativeText,owned=attachNativeFocusOwner({get:()=>manager} as unknown as Injector,api)
  const release=holdNativeTextFocus({...api}),releaseSecond=holdNativeTextFocus({...api})
  created$.next(render('late-draft',b))
  b.focus();b.activate();b.sync();b.addDocRanges()
  expect(nativeB).not.toHaveBeenCalled()
  owned.focus(()=>b.focus())
  // The actual native function was called only by explicit focus().
  expect(nativeB).toHaveBeenCalledOnce()
  release();release();b.focus();releaseSecond();b.focus()
  expect(nativeB).toHaveBeenCalledTimes(2)
  otherRuntime.focus();expect(otherRuntime.focus).toHaveBeenCalledOnce()
  disposed$.next('late-draft');expect(b.focus).toBe(nativeB)
  owned.dispose();expect(a.focus).toBe(original)
  expect(()=>holdNativeTextFocus(api)()).not.toThrow()
})
