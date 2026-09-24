import {afterEach,expect,it,vi} from 'vitest'
import type {Injector} from '@univerjs/core'
import {attachDerivedAutoHeight} from './nativeLayout'
afterEach(()=>vi.unstubAllGlobals())
it('re-measures sparse recovered rows in bounded, deduplicated local batches',()=>{
  const frames:(()=>void)[]=[],execute=vi.fn(),dispose=vi.fn()
  vi.stubGlobal('requestAnimationFrame',(f:()=>void)=>{frames.push(f);return frames.length})
  vi.stubGlobal('cancelAnimationFrame',vi.fn())
  const layout=attachDerivedAutoHeight({get:()=>({syncExecuteCommand:execute,onCommandExecuted:()=>({dispose})})} as unknown as Injector)
  layout.refresh('book','sheet',Array.from({length:260},(_,i)=>String(i*10)))
  layout.refresh('book','sheet',['0','10','-1','NaN'])
  while(frames.length)frames.shift()!()
  expect(execute.mock.calls.map(c=>c[1].ranges.length)).toEqual([128,128,4])
  expect(execute.mock.calls.every(c=>c[0]==='sheet.operation.mark-dirty-row-auto-height')).toBe(true)
  layout.dispose();expect(dispose).toHaveBeenCalledOnce()
})
