import {expect,it,vi} from 'vitest'
import {attachNativeRejectionReporter} from './nativeRejectionReporter'

type Listener = (command: unknown, options?: { fromCollab?: boolean; fromChangeset?: boolean }) => void

function harness() {
  let listener: Listener | null = null
  const events: Array<(event: { reason: unknown }) => void> = []
  const windowRef = {
    addEventListener: (_type: 'unhandledrejection', listener: (event: { reason: unknown }) => void) => { events.push(listener) },
    removeEventListener: (_type: 'unhandledrejection', listener: (event: { reason: unknown }) => void) => { const i = events.indexOf(listener); if (i >= 0) events.splice(i, 1) },
  }
  return {
    commands: { onCommandExecuted: (next: Listener) => { listener = next; return { dispose() { listener = null } } } },
    windowRef,
    events,
    runCommand(options?: { fromCollab?: boolean; fromChangeset?: boolean }) { listener?.({}, options) },
    reject(reason: unknown) { for (const fire of events) fire({ reason }) },
  }
}

it('reports an error rejection correlated with a recent local command, once per message',()=>{
  vi.useFakeTimers()
  try {
    vi.setSystemTime(0)
    const notify = vi.fn()
    const h = harness()
    const reporter = attachNativeRejectionReporter(h.commands, notify, h.windowRef)
    h.runCommand()
    h.reject(new TypeError("Cannot destructure property 'marginTop' of 'e.pages[(e.pages.length - 1)]' as it is undefined."))
    expect(notify).toHaveBeenCalledTimes(1)
    expect(notify.mock.calls[0][0]).toBeInstanceOf(Error)
    h.reject(new TypeError("Cannot destructure property 'marginTop' of 'e.pages[(e.pages.length - 1)]' as it is undefined."))
    expect(notify).toHaveBeenCalledTimes(1)
    h.reject(new TypeError('another failure'))
    expect(notify).toHaveBeenCalledTimes(2)
    reporter.dispose()
    h.reject(new TypeError('after dispose'))
    expect(notify).toHaveBeenCalledTimes(2)
    expect(h.events.length).toBe(0)
  } finally { vi.useRealTimers() }
})

it('ignores non-error reasons, remote commands, and rejections without a recent command',()=>{
  vi.useFakeTimers()
  try {
    vi.setSystemTime(0)
    const notify = vi.fn()
    const h = harness()
    const reporter = attachNativeRejectionReporter(h.commands, notify, h.windowRef)
    h.reject(new Error('no command ran yet'))
    expect(notify).not.toHaveBeenCalled()
    h.runCommand({ fromCollab: true })
    h.reject(new Error('remote echo'))
    expect(notify).not.toHaveBeenCalled()
    h.runCommand()
    vi.setSystemTime(3000)
    h.reject(new Error('correlation window elapsed'))
    expect(notify).not.toHaveBeenCalled()
    vi.setSystemTime(3001)
    h.runCommand()
    h.reject('not an Error instance')
    expect(notify).not.toHaveBeenCalled()
    reporter.dispose()
  } finally { vi.useRealTimers() }
})
