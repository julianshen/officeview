import { useMemo } from 'react'
import { OfficeFile } from 'officeview'

/** Download a file from a URL with a progress bar and a custom error message. */
export default function UrlWithProgress({ url }: { url: string }) {
  // Create the source once: a new object on every render would reload the file.
  const source = useMemo(() => ({ url }), [url])
  return (
    <OfficeFile
      data={source}
      loading={({ loaded, total }) => (
        <div style={{ margin: 'auto', width: 260, textAlign: 'center', font: '14px system-ui' }}>
          {total ? `Downloading ${Math.round((loaded / total) * 100)}%` : `Downloaded ${(loaded / 1024).toFixed(0)} KiB`}
          <div style={{ height: 4, marginTop: 8, background: '#e1dfdd', borderRadius: 2 }}>
            <div style={{ height: 4, width: total ? `${(loaded / total) * 100}%` : '30%', background: '#0f6cbd', borderRadius: 2 }} />
          </div>
        </div>
      )}
      error={(message) => (
        <p role="alert" style={{ margin: 'auto', color: '#c50f1f' }}>
          This document could not be opened ({message}).
        </p>
      )}
    />
  )
}
