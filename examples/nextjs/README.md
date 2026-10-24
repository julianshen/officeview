# officeview + Next.js

A minimal Next.js (App Router) app that opens a local Office file with
officeview, including the viewer worker for large spreadsheets. Tested with
Next.js 16 using both Turbopack (the default) and `next build --webpack`.

```bash
npm install
npm run dev
```

What to copy into your own app:

- `app/DocumentViewer.tsx` starts with `'use client'`: officeview uses React
  hooks and draws on `<canvas>`, so it renders in the browser. Server components
  (like `app/page.tsx`) can render the client component.
- `import 'officeview/style.css'` once, in a client component or a layout.
- `app/viewer-worker.ts` is a one-line worker entry, created with
  `new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })`.
  No `next.config` changes are needed.
