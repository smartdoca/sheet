import type { SpreadsheetAttachmentPreviewHandler, SpreadsheetNativeText } from './inlineTypes'

/** Keep callback updates independent of the native editor's lifecycle. */
export function subscribeAttachmentPreview(
  native: Pick<SpreadsheetNativeText, 'onNodeEvent'>,
  getHandler: () => SpreadsheetAttachmentPreviewHandler | undefined,
  onError: (error: Error) => void,
) {
  let disposed = false
  const reportError = (error: unknown) => {
    if (!disposed) onError(error instanceof Error ? error : new Error(String(error)))
  }
  const unsubscribe = native.onNodeEvent(event => {
    if (disposed || event?.phase !== 'click' || event.node.type !== 'attachment') return
    const handler = getHandler()
    if (!handler) return
    try {
      void Promise.resolve(handler({
        phase: 'click',
        node: { ...event.node, type: 'attachment' },
        cell: { ...event.cell },
      })).catch(reportError)
    } catch (error) {
      reportError(error)
    }
  })
  return () => { disposed = true; unsubscribe() }
}
