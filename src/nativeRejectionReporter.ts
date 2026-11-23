type Disposable = { dispose(): void }
type CommandSource = { fromCollab?: boolean; fromChangeset?: boolean }
type RejectionEvent = { reason: unknown }

/** Unhandled command-chain rejections (engine internals never await their own
 * executeCommand) look like dead gestures. Correlate a rejection with a recent
 * local command, report it once, and let the host toast it. */
export function attachNativeRejectionReporter(
  commands: { onCommandExecuted(listener: (command: unknown, options?: CommandSource) => void): Disposable },
  notify: (reason: Error) => void,
  windowRef: { addEventListener(type: 'unhandledrejection', listener: (event: RejectionEvent) => void): void; removeEventListener(type: 'unhandledrejection', listener: (event: RejectionEvent) => void): void } = window,
  correlationMs = 2000,
) {
  let lastCommandAt = -Infinity
  const offCommand = commands.onCommandExecuted((_command, options) => {
    if (options?.fromCollab || options?.fromChangeset) return
    lastCommandAt = Date.now()
  })
  const dedupe = new Map<string, number>()
  const onRejection = (event: RejectionEvent) => {
    const reason = event.reason
    if (!(reason instanceof Error)) return
    if (Date.now() - lastCommandAt > correlationMs) return
    const message = reason.message || String(reason)
    const last = dedupe.get(message)
    if (last !== undefined && Date.now() - last < 30000) return
    if (dedupe.size > 50) dedupe.clear()
    dedupe.set(message, Date.now())
    console.warn('[uos] unhandled rejection after a local sheet command', reason)
    notify(reason)
  }
  windowRef.addEventListener('unhandledrejection', onRejection)
  return { dispose() { offCommand.dispose(); windowRef.removeEventListener('unhandledrejection', onRejection) } }
}
