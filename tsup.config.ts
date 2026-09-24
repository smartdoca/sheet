import { defineConfig } from 'tsup'

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    yjs: 'src/yjs.ts',
    model: 'src/model.ts',
    xlsx: 'src/xlsx.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  sourcemap: true,
  clean: true,
  splitting: false,
  external: ['react', 'react-dom', 'yjs', 'y-protocols', 'y-indexeddb'],
  loader: { '.css': 'css' },
})
