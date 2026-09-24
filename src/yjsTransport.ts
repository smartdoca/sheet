import * as Y from 'yjs'
import {
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
  removeAwarenessStates,
  type Awareness,
} from 'y-protocols/awareness'

export type YjsConnectionStatus = 'offline' | 'syncing' | 'synced'

export type YjsSyncMessage =
  | { type: 'sync-request'; vector: Uint8Array }
  | { type: 'sync-response'; vector: Uint8Array; update: Uint8Array }
  | { type: 'update'; update: Uint8Array }
  | { type: 'awareness'; update: Uint8Array }

/** WebSocket/authentication/room routing stays in the host application. */
export interface YjsSyncTransport {
  send(message: YjsSyncMessage): void
  onMessage(callback: (message: YjsSyncMessage) => void): () => void
  onConnection(callback: (online: boolean) => void): () => void
  readonly connected: boolean
}

export interface ConnectYjsTransportOptions {
  /** Ephemeral member/cursor/selection state; never persisted by the backend. */
  awareness?: Awareness
}

/**
 * State-vector handshake. Edits made by either side while disconnected are
 * exchanged in both directions after reconnecting, without server-side OT.
 */
export function connectYjsTransport(
  doc: Y.Doc,
  transport: YjsSyncTransport,
  onStatus?: (status: YjsConnectionStatus) => void,
  options: ConnectYjsTransportOptions = {},
) {
  const remoteOrigin = { kind: 'univer-sheet-remote' }
  let online = transport.connected
  let destroyed = false
  const remoteAwarenessClients = new Set<number>()

  const localUpdate = (update: Uint8Array, origin: unknown) => {
    if (online && origin !== remoteOrigin) transport.send({ type: 'update', update })
  }
  const receive = (message: YjsSyncMessage) => {
    if (destroyed) return
    if (message.type === 'awareness') {
      if (options.awareness) applyAwarenessUpdate(options.awareness, message.update, remoteOrigin)
      return
    }
    if (message.type === 'sync-request') {
      transport.send({
        type: 'sync-response',
        update: Y.encodeStateAsUpdate(doc, message.vector),
        vector: Y.encodeStateVector(doc),
      })
      return
    }
    Y.applyUpdate(doc, message.update, remoteOrigin)
    if (message.type === 'sync-response') {
      transport.send({ type: 'update', update: Y.encodeStateAsUpdate(doc, message.vector) })
      onStatus?.('synced')
    }
  }
  const connection = (connected: boolean) => {
    online = connected
    onStatus?.(connected ? 'syncing' : 'offline')
    if (!connected && options.awareness && remoteAwarenessClients.size) {
      removeAwarenessStates(options.awareness, [...remoteAwarenessClients], remoteOrigin)
      remoteAwarenessClients.clear()
    }
    if (connected) {
      transport.send({ type: 'sync-request', vector: Y.encodeStateVector(doc) })
      if (options.awareness) {
        transport.send({
          type: 'awareness',
          update: encodeAwarenessUpdate(options.awareness, [options.awareness.clientID]),
        })
      }
    }
  }

  const awarenessUpdate = ({ added, updated, removed }: { added: number[]; updated: number[]; removed: number[] }, origin: unknown) => {
    if (!online || origin === remoteOrigin || !options.awareness) return
    transport.send({ type: 'awareness', update: encodeAwarenessUpdate(options.awareness, [...added, ...updated, ...removed]) })
  }
  const awarenessChange = ({ added, updated, removed }: { added: number[]; updated: number[]; removed: number[] }, origin: unknown) => {
    if (origin !== remoteOrigin) return
    added.forEach((id) => remoteAwarenessClients.add(id))
    updated.forEach((id) => remoteAwarenessClients.add(id))
    removed.forEach((id) => remoteAwarenessClients.delete(id))
  }

  const stopMessages = transport.onMessage(receive)
  const stopConnection = transport.onConnection(connection)
  doc.on('update', localUpdate)
  options.awareness?.on('update', awarenessUpdate)
  options.awareness?.on('change', awarenessChange)
  connection(online)

  return () => {
    destroyed = true
    stopMessages()
    stopConnection()
    doc.off('update', localUpdate)
    options.awareness?.off('update', awarenessUpdate)
    options.awareness?.off('change', awarenessChange)
  }
}
