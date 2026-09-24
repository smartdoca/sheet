import { expect, it } from 'vitest'
import * as Y from 'yjs'
import { Awareness } from 'y-protocols/awareness'

import {
  connectYjsTransport,
  type YjsSyncMessage,
  type YjsSyncTransport,
} from './yjsTransport'

it('exchanges edits made by both replicas while disconnected', () => {
  const documents = [new Y.Doc(), new Y.Doc()]
  const messages = [new Set<(message: YjsSyncMessage) => void>(), new Set<(message: YjsSyncMessage) => void>()]
  const connections = [new Set<(online: boolean) => void>(), new Set<(online: boolean) => void>()]
  let online = false

  const endpoint = (index: number): YjsSyncTransport => ({
    get connected() { return online },
    send(message) {
      if (online) messages[1 - index].forEach((receive) => receive(message))
    },
    onMessage(callback) {
      messages[index].add(callback)
      return () => { messages[index].delete(callback) }
    },
    onConnection(callback) {
      connections[index].add(callback)
      return () => { connections[index].delete(callback) }
    },
  })

  const stops = documents.map((document, index) => connectYjsTransport(document, endpoint(index)))
  documents[0].getText('cell').insert(0, 'A')
  documents[1].getText('cell').insert(0, 'B')
  online = true
  connections.forEach((callbacks) => callbacks.forEach((callback) => callback(true)))

  expect(documents[0].getText('cell').toString()).toEqual(documents[1].getText('cell').toString())
  expect(documents[0].getText('cell').length).toBe(2)
  stops.forEach((stop) => stop())
})

it('broadcasts awareness after reconnect without persisting it in the document', () => {
  const documents = [new Y.Doc(), new Y.Doc()]
  const awareness = documents.map((document) => new Awareness(document))
  const messages = [new Set<(message: YjsSyncMessage) => void>(), new Set<(message: YjsSyncMessage) => void>()]
  const connections = [new Set<(online: boolean) => void>(), new Set<(online: boolean) => void>()]
  let online = false
  const endpoint = (index: number): YjsSyncTransport => ({
    get connected() { return online },
    send(message) { if (online) messages[1 - index].forEach((receive) => receive(message)) },
    onMessage(callback) { messages[index].add(callback); return () => { messages[index].delete(callback) } },
    onConnection(callback) { connections[index].add(callback); return () => { connections[index].delete(callback) } },
  })
  const stops = documents.map((document, index) =>
    connectYjsTransport(document, endpoint(index), undefined, { awareness: awareness[index] }))
  awareness[0].setLocalState({ user: { id: 'alice', name: 'Alice' }, selection: 'A1:B2' })
  online = true
  connections.forEach((callbacks) => callbacks.forEach((callback) => callback(true)))

  expect(awareness[1].getStates().get(documents[0].clientID)).toEqual({
    user: { id: 'alice', name: 'Alice' }, selection: 'A1:B2',
  })
  expect(documents[0].share.size).toBe(0)
  stops.forEach((stop) => stop())
  awareness.forEach((state) => state.destroy())
})
