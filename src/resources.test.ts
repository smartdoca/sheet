import { expect, it, vi } from 'vitest'
import { ImageSourceType } from '@univerjs/core'
import { createUniverImageIoService } from './resources'

it('stores an asset ID and resolves the authorized URL only for current rendering', async () => {
  const resolve = vi.fn(async () => 'https://assets.example.test/temporary?token=ephemeral')
  const [, { useValue: service }] = createUniverImageIoService({
    upload: async () => ({ id: 'asset-uuid', kind: 'image', url: 'https://should-not-persist.example.test' }), resolve,
  }, 'book')
  const image = await service.saveImage(new File(['image'], 'image.png'))
  expect(image.source).toBe('asset-uuid')
  expect(image.imageSourceType).toBe(ImageSourceType.UUID)
  expect(await service.getImage(image.source)).toContain('temporary')
  expect(resolve).toHaveBeenCalledTimes(2)
})
