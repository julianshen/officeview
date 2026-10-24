import { useEffect, useState, type ReactNode } from 'react'
import { Icon } from './layout'
import './docs.css'

export type DocNav = Array<[group: string, items: Array<[id: string, label: string]>]>

export function Section({ id, title, children, lead }: { id: string; title: string; lead?: ReactNode; children: ReactNode }) {
  return (
    <section id={id} className="doc-section">
      <h2><a href={`#${id}`} className="anchor" aria-hidden="true">#</a>{title}</h2>
      {lead && <p className="doc-lead">{lead}</p>}
      {children}
    </section>
  )
}

export const H3 = ({ id, children }: { id?: string; children: ReactNode }) => <h3 id={id}>{children}</h3>

export type PropRow = [name: string, type: string, def: string, desc: ReactNode]

export function Props({ rows, head = 'Prop' }: { rows: PropRow[]; head?: string }) {
  return (
    <div className="table-wrap">
      <table className="data props">
        <thead><tr><th>{head}</th><th>Type</th><th>Default</th><th>Description</th></tr></thead>
        <tbody>
          {rows.map(([name, type, def, desc]) => (
            <tr key={name}>
              <td><code className="prop-name">{name}</code></td>
              <td><code className="prop-type">{type}</code></td>
              <td>{def ? <code>{def}</code> : <span className="muted">—</span>}</td>
              <td>{desc}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export function Table({ head, rows }: { head: string[]; rows: ReactNode[][] }) {
  return (
    <div className="table-wrap">
      <table className="data">
        <thead><tr>{head.map((h) => <th key={h}>{h}</th>)}</tr></thead>
        <tbody>{rows.map((r, i) => <tr key={i}>{r.map((cell, j) => <td key={j}>{cell}</td>)}</tr>)}</tbody>
      </table>
    </div>
  )
}

export const Note = ({ children, warn }: { children: ReactNode; warn?: boolean }) => (
  <div className={`callout${warn ? ' warn' : ''}`}>{warn ? <Icon.alert /> : <Icon.info />}<div>{children}</div></div>
)

export const c = (s: string) => <code>{s}</code>

/** Sidebar + content, with the current section highlighted as you scroll. */
export function DocShell({ nav, label, children }: { nav: DocNav; label: string; children: ReactNode }) {
  const [active, setActive] = useState(nav[0][1][0][0])
  const [open, setOpen] = useState(false)
  useEffect(() => {
    const sections = Array.from(document.querySelectorAll<HTMLElement>('.doc-section'))
    const io = new IntersectionObserver(
      (entries) => {
        const visible = entries.filter((e) => e.isIntersecting).sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top)
        if (visible[0]) setActive(visible[0].target.id)
      },
      { rootMargin: '-80px 0px -65% 0px' },
    )
    sections.forEach((s) => io.observe(s))
    // The content renders after the browser has handled the URL fragment.
    if (location.hash) document.getElementById(location.hash.slice(1))?.scrollIntoView()
    return () => io.disconnect()
  }, [])
  return (
    <main className="container docs-layout">
      <button className="docs-nav-toggle btn btn-secondary btn-sm" onClick={() => setOpen((o) => !o)} aria-expanded={open}>
        <Icon.menu /> Contents
      </button>
      <nav className={`docs-nav${open ? ' open' : ''}`} aria-label={label}>
        {nav.map(([group, items]) => (
          <div key={group} className="docs-nav-group">
            <div className="docs-nav-title">{group}</div>
            {items.map(([id, text]) => (
              <a key={id} href={`#${id}`} className={active === id ? 'active' : ''} onClick={() => setOpen(false)}>{text}</a>
            ))}
          </div>
        ))}
      </nav>
      <article className="doc-content">{children}</article>
    </main>
  )
}
