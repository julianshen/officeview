'use client'

// officeview renders on <canvas> with React hooks, so it runs in the browser:
// keep it in a client component.
import { useState } from 'react'
import { OfficeFile } from 'officeview'
import 'officeview/style.css'

const createViewerWorker = () => new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })

export function DocumentViewer() {
  const [file, setFile] = useState<File | null>(null)
  return (
    <>
      <header style={{ display: 'flex', gap: 12, alignItems: 'center', padding: 12, borderBottom: '1px solid #ddd' }}>
        <strong>officeview + Next.js</strong>
        <input type="file" accept=".docx,.xlsx,.pptx,.odt,.rtf" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
      </header>
      {/* The viewer scrolls its own pages, so it sits in a sized flex column. */}
      {file ? (
        <OfficeFile data={file} interactiveSheets viewerWorker={createViewerWorker} />
      ) : (
        <p style={{ padding: 24, color: '#666' }}>Choose a .docx, .xlsx, .pptx, .odt or .rtf file.</p>
      )}
    </>
  )
}
