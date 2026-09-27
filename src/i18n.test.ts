import { expect, it } from 'vitest'
import {
  createTranslator,
  enEditorMessages,
  resolveLocale,
  zhEditorMessages,
} from './i18n'

it('keeps Chinese and English catalogs on the same keys', () => {
  expect(Object.keys(zhEditorMessages).sort()).toEqual(Object.keys(enEditorMessages).sort())
  expect(zhEditorMessages['locale.zh']).toBe('中文')
  expect(enEditorMessages['locale.zh']).toBe('中文')
  expect(zhEditorMessages['locale.en']).toBe('English')
  expect(enEditorMessages['locale.en']).toBe('English')
})

it('resolves omitted and Chinese codes to zh, and everything else to en', () => {
  expect(resolveLocale()).toBe('zh')
  expect(resolveLocale('')).toBe('zh')
  expect(resolveLocale('zh')).toBe('zh')
  expect(resolveLocale('zh-CN')).toBe('zh')
  expect(resolveLocale('zh_Hans')).toBe('zh')
  expect(resolveLocale('en')).toBe('en')
  expect(resolveLocale('en-US')).toBe('en')
  expect(resolveLocale('fr')).toBe('en')
  expect(createTranslator()('action.save')).toBe('保存')
  expect(createTranslator('zh-CN')('action.save')).toBe('保存')
  expect(createTranslator('en-US')('action.save')).toBe('Save')
  expect(createTranslator('fr')('action.save')).toBe('Save')
})

it('fills placeholders, selects plurals, and overrides individual keys', () => {
  expect(createTranslator('zh')('sheet.deleteConfirm', { name: 'Q1' })).toBe('确定删除工作表“Q1”吗？')
  expect(createTranslator('en')('chart.summary', { title: 'Sales', count: 1 })).toBe('Sales, 1 record')
  expect(createTranslator('en')('chart.summary', { title: 'Sales', count: 0 })).toBe('Sales, 0 records')
  expect(createTranslator('zh')('chart.summary', { title: '销售', count: 2 })).toBe('销售，2 条数据')
  expect(createTranslator('en', { 'action.save': 'Store' })('action.save')).toBe('Store')
  expect(createTranslator('en', { 'action.save': 'Store' })('action.undo')).toBe('Undo')
})

it('falls back to English and then to the key', () => {
  const saved = zhEditorMessages['action.save']
  delete zhEditorMessages['action.save']
  try {
    expect(createTranslator('zh')('action.save')).toBe('Save')
    expect(createTranslator('zh')('missing.key')).toBe('missing.key')
  } finally {
    zhEditorMessages['action.save'] = saved
  }
})
