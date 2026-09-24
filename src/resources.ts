import { IImageIoService, ImageSourceType, ImageUploadStatusType } from '@univerjs/core'
import { Subject } from 'rxjs'

import type { ResourceAdapter, SpreadsheetResource } from './types'

export function createUniverImageIoService(adapter: ResourceAdapter, workbookId: string) {
  const changes = new Subject<number>()
  const cache = new Map<string, HTMLImageElement>()
  let waitCount = 0

  const service = {
    change$: changes.asObservable(),
    setWaitCount(count: number) {
      waitCount = count
      changes.next(waitCount)
    },
    async getImage(imageId: string) {
      return adapter.resolve({ id: imageId, kind: 'image' }, { workbookId })
    },
    async saveImage(file: File) {
      try {
        const resource = await adapter.upload(file, { kind: 'image' }, { workbookId })
        if (!resource.id) throw new Error('Resource upload must return a stable resource ID')
        const resolved = await adapter.resolve(resource, { workbookId })
        return {
          imageId: resource.id,
          imageSourceType: ImageSourceType.UUID,
          source: resource.id,
          base64Cache: resolved,
          status: ImageUploadStatusType.SUCCUSS,
        }
      } finally {
        waitCount = Math.max(0, waitCount - 1)
        changes.next(waitCount)
      }
    },
    getImageSourceCache(source: string, sourceType: ImageSourceType) {
      if (sourceType === ImageSourceType.BASE64) {
        const image = new Image()
        image.src = source
        return image
      }
      return cache.get(source) ?? null
    },
    addImageSourceCache(source: string, sourceType: ImageSourceType, image: HTMLImageElement | null) {
      if (sourceType !== ImageSourceType.BASE64 && image) cache.set(source, image)
    },
  }

  return [IImageIoService, { useValue: service }] as [typeof IImageIoService, { useValue: typeof service }]
}

export async function downloadResourceResult(result: Blob | string | void, resource: SpreadsheetResource) {
  if (!result) return
  const url = typeof result === 'string' ? result : URL.createObjectURL(result)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = resource.name ?? resource.id
  anchor.click()
  if (result instanceof Blob) URL.revokeObjectURL(url)
}
