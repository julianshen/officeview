import { useMemo, useRef, useState } from 'react'
import { OfficeFile, SlideShowError, type OfficeViewerHandle } from 'officeview'

/** Start a slideshow from your own button and follow it with the slideshow callbacks. */
export default function PresentControls({ url }: { url: string }) {
  // Create the source once: a new object on every render would reload the file.
  const source = useMemo(() => ({ url }), [url])
  const viewer = useRef<OfficeViewerHandle>(null)
  const [status, setStatus] = useState('Not presenting')

  const present = (from: 'beginning' | number) => {
    // Call present() synchronously in the click handler: browsers only allow
    // full screen during a user gesture.
    viewer.current?.present({ from })
      .then((session) => session.ended)
      .catch((e) => setStatus(e instanceof SlideShowError ? `Could not present: ${e.code}` : String(e)))
  }

  return (
    <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', padding: 12, borderBottom: '1px solid #e1dfdd', font: '14px system-ui' }}>
        <button onClick={() => present('beginning')}>Present</button>
        <button onClick={() => present(2)}>Present from slide 3</button>
        <span style={{ color: '#605e5c' }}>{status}</span>
      </div>
      <OfficeFile
        ref={viewer}
        data={source}
        showToolbar={false}
        onSlideShowStart={() => setStatus('Presenting')}
        onSlideShowSlideChange={(slide, build) => setStatus(`Slide ${slide + 1}, step ${build}`)}
        onSlideShowEnd={() => setStatus('Show ended')}
      />
    </div>
  )
}
