import { useEffect, useState } from 'react'
import { buildTextIndex, findMatches, getPaintables, loadOfficeFile, type TextIndex } from 'officeview'

/** Search a document without displaying it: build the same text index the viewer uses. */
export default function TextSearch({ url }: { url: string }) {
  const [index, setIndex] = useState<TextIndex | null>(null)
  const [query, setQuery] = useState('revenue')
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false
    loadOfficeFile({ url })
      .then((doc) => getPaintables(doc))
      .then(async (pages) => {
        try {
          const built = await buildTextIndex(pages)
          if (!cancelled) setIndex(built)
        } finally {
          pages.dispose()   // release fonts and images even if indexing fails
        }
      })
      .catch((e) => { if (!cancelled) setError(e instanceof Error ? e.message : String(e)) })
    return () => { cancelled = true }
  }, [url])

  const matches = index && query.trim() ? findMatches(index, query) : []

  return (
    <div style={{ padding: 16, overflow: 'auto', font: '14px/1.5 system-ui' }}>
      <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search…" style={{ width: '100%', padding: 8, fontSize: 14 }} />
      <p style={{ color: '#605e5c' }}>{error ? `Could not index the document: ${error}` : index ? `${matches.length} matches` : 'Indexing…'}</p>
      <ol style={{ paddingLeft: 20 }}>
        {matches.slice(0, 50).map((m, i) => {
          const line = index!.pages[m.pageIndex].lines[m.lineIndex].text
          return (
            <li key={i} style={{ marginBottom: 6 }}>
              <span style={{ color: '#605e5c' }}>Page {m.pageIndex + 1}: </span>
              {line.slice(Math.max(0, m.start - 40), m.start)}
              <mark>{line.slice(m.start, m.end)}</mark>
              {line.slice(m.end, m.end + 40)}
            </li>
          )
        })}
      </ol>
    </div>
  )
}
