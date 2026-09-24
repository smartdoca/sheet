import { expect, it } from 'vitest'
import { BorderStyleTypes, BorderType, LocaleType, Univer, UniverInstanceType, type Workbook } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { FUniver } from '@univerjs/core/facade'
import '@univerjs/sheets/facade'

it('keeps explicit border color/style independent from fill, including undo/redo', async () => {
  const univer = new Univer({ locale: LocaleType.EN_US, locales: { [LocaleType.EN_US]: {} } })
  univer.registerPlugin(UniverSheetsPlugin)
  try {
    const model = univer.createUnit(UniverInstanceType.UNIVER_SHEET, { id: 'borders' }) as Workbook
    const api = FUniver.newAPI(univer)
    const range = api.getActiveWorkbook()!.getActiveSheet()!.getRange('A1:B2')
    const style = () => model.getStyles().getStyleByCell(model.getActiveSheet()!.getCell(0, 0)!)!
    range.setBorder(BorderType.ALL, BorderStyleTypes.MEDIUM, '#ef4444')
    const border = structuredClone(style().bd)
    range.setBackgroundColor('#dbeafe')
    expect(style().bd).toEqual(border)
    expect(style().bg?.rgb).toBe('#dbeafe')
    range.setBackgroundColor('#fef3c7')
    expect(style().bd).toEqual(border)
    api.getActiveWorkbook()!.undo()
    expect(style().bg?.rgb).toBe('#dbeafe')
    expect(style().bd).toEqual(border)
    api.getActiveWorkbook()!.redo()
    expect(style().bg?.rgb).toBe('#fef3c7')
    expect(style().bd).toEqual(border)
    range.setBorder(BorderType.ALL, BorderStyleTypes.THIN, '#2563eb')
    expect(style().bg?.rgb).toBe('#fef3c7')
    expect(style().bd?.t?.cl?.rgb).toBe('#2563eb')
    range.setBackgroundColor(null as unknown as string)
    expect(style().bd?.t?.cl?.rgb).toBe('#2563eb')
  } finally { univer.dispose() }
})
