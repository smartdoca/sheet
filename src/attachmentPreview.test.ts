import { describe, expect, it, vi } from 'vitest'
import { subscribeAttachmentPreview } from './attachmentPreview'
import type { SpreadsheetAttachmentPreviewHandler, SpreadsheetInlineNodeEvent } from './inlineTypes'

const click: SpreadsheetInlineNodeEvent = {
  phase: 'click',
  node: { type: 'attachment', refId: 'asset-1', label: '📎 说明.pdf' },
  cell: { workbookId: 'isolated-book', sheetId: 's', row: 2, column: 3 },
}

function source() {
  const listeners = new Set<(event: SpreadsheetInlineNodeEvent | null) => void>()
  return {
    native: { onNodeEvent(listener: (event: SpreadsheetInlineNodeEvent | null) => void) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    } },
    emit(event: SpreadsheetInlineNodeEvent | null) { listeners.forEach(listener => listener(event)) },
    listeners,
  }
}

describe('host attachment preview', () => {
  it('only previews attachment clicks and gives the host an isolated identity/position payload', () => {
    const events = source(), onError = vi.fn()
    const handler = vi.fn<SpreadsheetAttachmentPreviewHandler>(event => {
      event.node.label = 'host display name'
      event.cell.row = 99
    })
    const dispose = subscribeAttachmentPreview(events.native, () => handler, onError)
    events.emit(null)
    events.emit({ ...click, phase: 'hover' })
    for (const type of ['user', 'document', 'image']) events.emit({ ...click, node: { ...click.node, type } })
    expect(handler).not.toHaveBeenCalled()
    const event = structuredClone(click)
    events.emit(event)
    expect(handler).toHaveBeenCalledTimes(1)
    expect(handler.mock.calls[0][0].node.refId).toBe('asset-1')
    expect(handler.mock.calls[0][0].phase).toBe('click')
    expect(event).toEqual(click)
    expect(onError).not.toHaveBeenCalled()
    dispose()
    expect(events.listeners.size).toBe(0)
    events.emit(click)
    expect(handler).toHaveBeenCalledTimes(1)
  })

  it('uses the latest optional handler without rebuilding the subscription', () => {
    const events = source(), first = vi.fn(), second = vi.fn()
    let handler: SpreadsheetAttachmentPreviewHandler | undefined
    const dispose = subscribeAttachmentPreview(events.native, () => handler, vi.fn())
    events.emit(click)
    handler = first; events.emit(click)
    handler = second; events.emit(click)
    handler = undefined; events.emit(click)
    expect(first).toHaveBeenCalledExactlyOnceWith(click)
    expect(second).toHaveBeenCalledExactlyOnceWith(click)
    expect(events.listeners.size).toBe(1)
    dispose()
  })

  it('reports synchronous and asynchronous failures as errors', async () => {
    const events = source(), onError = vi.fn(), denied = new Error('Asset access denied')
    let handler: SpreadsheetAttachmentPreviewHandler = () => { throw denied }
    const dispose = subscribeAttachmentPreview(events.native, () => handler, onError)
    events.emit(click)
    expect(onError).toHaveBeenLastCalledWith(denied)
    handler = () => Promise.reject('Preview unavailable')
    events.emit(click)
    await Promise.resolve()
    expect(onError).toHaveBeenLastCalledWith(new Error('Preview unavailable'))
    expect(onError).toHaveBeenCalledTimes(2)
    dispose()
  })

  it('ignores pending preview failures after disposal', async () => {
    const events = source(), onError = vi.fn()
    let reject!: (error: Error) => void
    const pending = new Promise<void>((_, fail) => { reject = fail })
    const dispose = subscribeAttachmentPreview(events.native, () => () => pending, onError)
    events.emit(click)
    dispose()
    reject(new Error('Preview closed'))
    await Promise.resolve()
    expect(onError).not.toHaveBeenCalled()
  })
})
