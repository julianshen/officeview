import { useState, type DragEvent } from 'react'
import { OfficeFile } from 'officeview'

/** Open a file from the user's device: a picker plus drag and drop. Nothing is uploaded. */
export default function FilePicker() {
  const [file, setFile] = useState<File | null>(null)
  const [over, setOver] = useState(false)

  const onDrop = (e: DragEvent) => {
    e.preventDefault()
    setOver(false)
    const dropped = e.dataTransfer.files[0]
    if (dropped) setFile(dropped)
  }

  return (
    <div
      style={{ height: '100%', display: 'flex', flexDirection: 'column', outline: over ? '2px dashed #0f6cbd' : 'none' }}
      onDragOver={(e) => { e.preventDefault(); setOver(true) }}
      onDragLeave={() => setOver(false)}
      onDrop={onDrop}
    >
      <div style={{ padding: 12, borderBottom: '1px solid #e1dfdd' }}>
        <input type="file" accept=".docx,.xlsx,.pptx,.odt,.rtf" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
      </div>
      {file ? (
        <OfficeFile data={file} />
      ) : (
        <p style={{ margin: 'auto', color: '#605e5c' }}>Choose a file, or drop one here.</p>
      )}
    </div>
  )
}
