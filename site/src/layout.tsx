import { StrictMode, useEffect, useState, type ReactNode } from 'react'
import { createRoot } from 'react-dom/client'
import pkg from '../../package.json'
import './site.css'

export const BASE = import.meta.env.BASE_URL
export const VERSION = pkg.version
export const REPO = 'https://github.com/julianshen/officeview'
export const NPM = 'https://www.npmjs.com/package/officeview'

export const href = (path: string) => `${BASE}${path.replace(/^\//, '')}`

export type SiteTheme = 'light' | 'dark'

function systemTheme(): SiteTheme {
  return typeof matchMedia === 'function' && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

function storedTheme(): SiteTheme | null {
  try {
    const v = localStorage.getItem('ov-site-theme')
    return v === 'light' || v === 'dark' ? v : null
  } catch {
    return null
  }
}

/** The page's effective colour theme; viewers on the page follow it. */
export function useSiteTheme(): [SiteTheme, () => void] {
  const [theme, setTheme] = useState<SiteTheme>(() => storedTheme() ?? systemTheme())
  useEffect(() => {
    if (storedTheme()) return
    const mq = matchMedia('(prefers-color-scheme: dark)')
    const on = () => setTheme(mq.matches ? 'dark' : 'light')
    mq.addEventListener('change', on)
    return () => mq.removeEventListener('change', on)
  }, [])
  useEffect(() => {
    document.documentElement.dataset.theme = theme
    const handler = (e: Event) => setTheme((e as CustomEvent<SiteTheme>).detail)
    window.addEventListener('ov-site-theme', handler)
    return () => window.removeEventListener('ov-site-theme', handler)
  }, [theme])
  const toggle = () => {
    const next = theme === 'dark' ? 'light' : 'dark'
    try { localStorage.setItem('ov-site-theme', next) } catch { /* private mode */ }
    window.dispatchEvent(new CustomEvent('ov-site-theme', { detail: next }))
  }
  return [theme, toggle]
}

const P = { fill: 'none', stroke: 'currentColor', strokeWidth: 2, strokeLinecap: 'round', strokeLinejoin: 'round' } as const

export const Icon = {
  github: () => <svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 .5a11.5 11.5 0 0 0-3.64 22.41c.58.1.79-.25.79-.56v-2c-3.2.7-3.88-1.37-3.88-1.37-.52-1.33-1.28-1.69-1.28-1.69-1.05-.71.08-.7.08-.7 1.16.08 1.77 1.19 1.77 1.19 1.03 1.77 2.7 1.26 3.36.96.1-.75.4-1.26.73-1.55-2.56-.29-5.25-1.28-5.25-5.69 0-1.26.45-2.29 1.19-3.1-.12-.29-.52-1.46.11-3.05 0 0 .97-.31 3.17 1.18a11 11 0 0 1 5.77 0c2.2-1.49 3.17-1.18 3.17-1.18.63 1.59.23 2.76.11 3.05.74.81 1.19 1.84 1.19 3.1 0 4.42-2.7 5.39-5.26 5.68.41.36.78 1.06.78 2.14v3.17c0 .31.21.67.8.56A11.5 11.5 0 0 0 12 .5Z" /></svg>,
  sun: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><circle cx="12" cy="12" r="4" /><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41" /></svg>,
  moon: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79Z" /></svg>,
  menu: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M4 6h16M4 12h16M4 18h16" /></svg>,
  copy: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><rect x="9" y="9" width="13" height="13" rx="2" /><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" /></svg>,
  check: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M20 6 9 17l-5-5" /></svg>,
  arrow: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M5 12h14M13 6l6 6-6 6" /></svg>,
  download: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M12 3v12M7 10l5 5 5-5M5 21h14" /></svg>,
  upload: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M12 21V9M7 14l5-5 5 5M5 3h14" /></svg>,
  info: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><circle cx="12" cy="12" r="10" /><path d="M12 16v-4M12 8h.01" /></svg>,
  alert: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z" /><path d="M12 9v4M12 17h.01" /></svg>,
  play: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="m6 4 14 8-14 8V4Z" /></svg>,
  canvas: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><rect x="3" y="3" width="18" height="18" rx="2" /><path d="M3 9h18M9 21V9" /></svg>,
  layers: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="m12 2 10 5-10 5L2 7l10-5Z" /><path d="m2 17 10 5 10-5M2 12l10 5 10-5" /></svg>,
  chart: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M3 3v18h18" /><path d="M7 15v3M12 10v8M17 6v12" /></svg>,
  present: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><rect x="2" y="3" width="20" height="13" rx="2" /><path d="M12 16v5M8 21h8" /></svg>,
  stream: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M21 12a9 9 0 1 1-6.22-8.56" /><path d="M21 3v6h-6" /></svg>,
  shield: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z" /></svg>,
  palette: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><circle cx="13.5" cy="6.5" r="1.5" /><circle cx="17.5" cy="10.5" r="1.5" /><circle cx="8.5" cy="7.5" r="1.5" /><circle cx="6.5" cy="12.5" r="1.5" /><path d="M12 2a10 10 0 0 0 0 20c.93 0 1.5-.75 1.5-1.5 0-.39-.15-.74-.39-1.01-.24-.27-.38-.62-.38-1.01 0-.83.67-1.5 1.5-1.5H16a6 6 0 0 0 6-6c0-4.97-4.48-9-10-9Z" /></svg>,
  grid: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><rect x="3" y="3" width="18" height="18" rx="2" /><path d="M3 9h18M3 15h18M9 3v18M15 3v18" /></svg>,
  search: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><circle cx="11" cy="11" r="7" /><path d="m21 21-4.3-4.3" /></svg>,
  phone: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><rect x="6" y="2" width="12" height="20" rx="2" /><path d="M11 18h2" /></svg>,
  stamp: () => <svg viewBox="0 0 24 24" aria-hidden="true" {...P}><path d="M5 22h14M5 18h14v-3a3 3 0 0 0-3-3h-2V8a3 3 0 1 0-4 0v4H8a3 3 0 0 0-3 3v3Z" /></svg>,
}

export function Logo() {
  return (
    <svg viewBox="0 0 32 32" aria-hidden="true">
      <rect x="2" y="2" width="28" height="28" rx="7" fill="#0f6cbd" />
      <rect x="8" y="8" width="11" height="14" rx="1.5" fill="#fff" />
      <path d="M10.5 12h6M10.5 15h6M10.5 18h4" stroke="#0f6cbd" strokeWidth="1.4" strokeLinecap="round" />
      <rect x="15" y="13" width="9" height="11" rx="1.5" fill="#7cc0ff" stroke="#0f6cbd" strokeWidth="1.2" />
    </svg>
  )
}

type PageId = 'home' | 'demos' | 'playground' | 'docs'

const NAV: Array<[PageId, string, string]> = [
  ['demos', 'Demos', 'demos/'],
  ['playground', 'Playground', 'playground/'],
  ['docs', 'Docs', 'docs/'],
]

export function Header({ page }: { page: PageId }) {
  const [theme, toggle] = useSiteTheme()
  const [open, setOpen] = useState(false)
  return (
    <header className="site-header">
      <div className="container">
        <a className="brand" href={href('')}>
          <Logo />
          officeview
          <span className="version">v{VERSION}</span>
        </a>
        <nav className={`nav${open ? ' open' : ''}`} aria-label="Main">
          {NAV.map(([id, label, path]) => (
            <a key={id} href={href(path)} aria-current={page === id ? 'page' : undefined}>{label}</a>
          ))}
        </nav>
        <div className="header-actions">
          <button className="icon-btn" onClick={toggle} aria-label={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'} title="Toggle theme">
            {theme === 'dark' ? <Icon.sun /> : <Icon.moon />}
          </button>
          <a className="icon-btn" href={REPO} aria-label="GitHub repository" title="GitHub"><Icon.github /></a>
          <button className="icon-btn menu-btn" onClick={() => setOpen((o) => !o)} aria-label="Menu" aria-expanded={open}><Icon.menu /></button>
        </div>
      </div>
    </header>
  )
}

export function Footer() {
  return (
    <footer className="site-footer">
      <div className="container">
        <div>officeview v{VERSION} · MIT licensed · Sample documents are fictional.</div>
        <nav aria-label="Footer">
          <a href={href('docs/')}>Documentation</a>
          <a href={REPO}>GitHub</a>
          <a href={NPM}>npm</a>
          <a href={`${REPO}/blob/main/CHANGELOG.md`}>Changelog</a>
          <a href={`${REPO}/issues`}>Issues</a>
        </nav>
      </div>
    </footer>
  )
}

export function mount(page: PageId, content: ReactNode, opts: { footer?: boolean } = {}) {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <Header page={page} />
      {content}
      {opts.footer !== false && <Footer />}
    </StrictMode>,
  )
}
