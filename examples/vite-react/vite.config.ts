import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // officeview's viewer worker code-splits its lazy modules, which needs
  // ES-module worker output (Vite's default worker format is 'iife').
  worker: { format: 'es' },
})
