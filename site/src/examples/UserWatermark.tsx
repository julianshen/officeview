import { useMemo, useState } from 'react'
import { OfficeFile } from 'officeview'

/** Stamp every page with who is viewing it and when, and turn off copying and printing. */
export default function UserWatermark({ url }: { url: string }) {
  // Create the source once: a new object on every render would reload the file.
  const source = useMemo(() => ({ url }), [url])
  const [user, setUser] = useState('jordan@example.com')
  const stamp = `${user} · ${new Date().toLocaleDateString()}`
  return (
    <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <label style={{ display: 'flex', gap: 8, alignItems: 'center', padding: 12, borderBottom: '1px solid #e1dfdd', font: '14px system-ui' }}>
        Viewer
        <input value={user} onChange={(e) => setUser(e.target.value)} style={{ padding: 6, width: 240 }} />
      </label>
      <OfficeFile
        data={source}
        watermark={{ text: stamp, placement: 'tile', rotate: -30, opacity: 0.1, fontSizePt: 18 }}
        allowCopy={false}
        allowPrint={false}
      />
    </div>
  )
}
