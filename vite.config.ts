import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  plugins: [react()],
  resolve:{alias:[
    {find:/^@online-office\/univer-sheet\/model$/,replacement:new URL('./src/model.ts',import.meta.url).pathname},
    {find:/^@online-office\/univer-sheet\/xlsx$/,replacement:new URL('./src/xlsx.ts',import.meta.url).pathname},
    {find:/^@online-office\/univer-sheet\/yjs$/,replacement:new URL('./src/yjs.ts',import.meta.url).pathname},
    {find:/^@online-office\/univer-sheet\/style.css$/,replacement:new URL('./src/style.css',import.meta.url).pathname},
    {find:/^@online-office\/univer-sheet$/,replacement:new URL('./src/index.ts',import.meta.url).pathname},
  ]},
  build: {
    outDir: 'dist-demo',
  },
})
