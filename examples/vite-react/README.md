# officeview + Vite

A minimal React app that opens a local Office file with officeview, including
the viewer worker for large spreadsheets.

```bash
npm install
npm run dev
```

The two things to copy into your own app:

- `vite.config.ts` sets `worker: { format: 'es' }`.
- `src/viewer-worker.ts` is a one-line worker entry, created in `src/App.tsx`
  with `new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })`.
