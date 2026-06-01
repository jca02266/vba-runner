import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import path from 'path'

// Web UI (Playground) のビルド設定。
// ソース（index.html / src/App.tsx 等）はリポジトリルートに据え置きのため
// root をルートに向け、ビルド出力だけを build/playground/dist に隔離する。
const repoRoot = path.resolve(import.meta.dirname, '../../')

export default defineConfig({
  root: repoRoot,
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      path: 'path-browserify',
    },
  },
  build: {
    outDir: path.resolve(import.meta.dirname, 'dist'),
    emptyOutDir: true,
  },
})
