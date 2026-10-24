import { useMemo } from 'react'
import { OfficeDoc, useOfficeFile } from 'officeview'

/** Own the loading, error and notice UI with useOfficeFile, and render the result with <OfficeDoc>. */
export default function CustomStates({ url }: { url: string }) {
  // Create the source once: a new object on every render would reload the file.
  const source = useMemo(() => ({ url }), [url])
  const state = useOfficeFile(source)

  if (state.status === 'loading') {
    return <p style={{ margin: 'auto', font: '14px system-ui', color: '#605e5c' }}>Preparing your document…</p>
  }
  if (state.status === 'error') {
    return (
      <div role="alert" style={{ margin: 'auto', font: '14px system-ui', color: '#c50f1f' }}>
        <strong>Could not open the document.</strong> {state.message}
        {state.code && <div>Code: {state.code}</div>}
      </div>
    )
  }
  return (
    <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      {/* Notices (for example "Showing saved values") are yours to display. */}
      {state.notices?.map((notice) => (
        <div key={notice} role="status" style={{ padding: '8px 12px', background: '#fff4ce', font: '13px system-ui' }}>{notice}</div>
      ))}
      <OfficeDoc document={state.document} title="Operating budget" />
    </div>
  )
}
