import { useMemo, useState } from 'react'
import { OfficeFile } from 'officeview'

// Re-brand the viewer chrome by overriding its --ov-* tokens on the theme root.
const css = `
.brand-red [data-ov-theme-root] {
  --ov-color-accent: #c50f1f;
  --ov-color-chrome: #fdf3f4;
  --ov-color-backdrop: #f6e9eb;
  --ov-color-border: #efc6cb;
}
.brand-red [data-ov-theme-root][data-ov-theme="dark"] {
  --ov-color-accent: #ff99a4;
  --ov-color-chrome: #2b1d1f;
  --ov-color-backdrop: #1d1415;
  --ov-color-border: #4a3034;
}`

export default function BrandTheme({ url }: { url: string }) {
  // Create the source once: a new object on every render would reload the file.
  const source = useMemo(() => ({ url }), [url])
  const [theme, setTheme] = useState<'light' | 'dark'>('light')
  return (
    <div className="brand-red" style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <style>{css}</style>
      <div style={{ padding: 12, borderBottom: '1px solid #e1dfdd', font: '14px system-ui' }}>
        <button onClick={() => setTheme(theme === 'light' ? 'dark' : 'light')}>Switch to {theme === 'light' ? 'dark' : 'light'}</button>
      </div>
      <OfficeFile data={source} theme={theme} />
    </div>
  )
}
