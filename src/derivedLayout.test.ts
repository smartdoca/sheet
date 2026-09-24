import {expect,it} from 'vitest'
import {isDerivedLayoutCommand} from './derivedLayout'
it('only exempts derived measurements, not authored row sizes or flags',()=>{
  expect(isDerivedLayoutCommand('sheet.mutation.set-worksheet-row-auto-height')).toBe(true)
  for(const id of ['sheet.command.set-row-height','sheet.mutation.set-worksheet-row-height','sheet.mutation.set-worksheet-row-is-auto-height','sheet.mutation.set-worksheet-col-width'])expect(isDerivedLayoutCommand(id)).toBe(false)
})
