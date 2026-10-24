import { useState, type DragEvent } from 'react'
import { OfficeFile } from 'officeview'
import 'officeview/style.css'

const createViewerWorker = () => new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })

export function App() {
  const [file, setFile] = useState<File | null>(null)

  const onDrop = (e: DragEvent) => {
    e.preventDefault()
    const dropped = e.dataTransfer.files[0]
    if (dropped) setFile(dropped)
  }

  return (
    // The viewer scrolls its own pages, so it needs a sized container.
    <div
      style={{ height: '100vh', display: 'flex', flexDirection: 'column', fontFamily: 'system-ui, sans-serif' }}
      onDragOver={(e) => e.preventDefault()}
      onDrop={onDrop}
    >
      <header style={{ display: 'flex', gap: 12, alignItems: 'center', padding: 12, borderBottom: '1px solid #ddd' }}>
        <strong>officeview + Vite</strong>
        <input type="file" accept=".docx,.xlsx,.pptx,.odt,.rtf" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
        <span style={{ color: '#666' }}>or drop a file anywhere</span>
      </header>
      {file ? (
        <OfficeFile
          data={file}
          interactiveSheets
          viewerWorker={createViewerWorker}
          loading={({ loaded, total }) => <p style={{ padding: 24 }}>Loading… {total ? `${Math.round((loaded / total) * 100)}%` : ''}</p>}
          error={(message) => <p style={{ padding: 24, color: '#c00' }}>Could not open this file: {message}</p>}
        />
      ) : (
        <p style={{ padding: 24, color: '#666' }}>Choose a .docx, .xlsx, .pptx, .odt or .rtf file.</p>
      )}
    </div>
  )
}
