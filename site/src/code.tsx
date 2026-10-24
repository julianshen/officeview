import { useState, type ReactNode } from 'react'
import { Icon } from './layout'

export type Lang = 'tsx' | 'ts' | 'css' | 'bash' | 'http' | 'html'

const KEYWORDS = new Set([
  'import', 'from', 'export', 'default', 'const', 'let', 'var', 'function', 'return', 'await', 'async', 'new', 'if', 'else',
  'type', 'interface', 'extends', 'as', 'typeof', 'null', 'undefined', 'true', 'false', 'void', 'catch', 'try', 'for', 'of', 'in',
])

/**
 * A small single-pass highlighter: enough for the snippets on this site
 * (TS/TSX, CSS, shell and HTTP headers), with no runtime dependency.
 */
function highlight(code: string, lang: Lang): ReactNode[] {
  const out: ReactNode[] = []
  let key = 0
  const push = (cls: string | null, text: string) => out.push(cls ? <span key={key++} className={cls}>{text}</span> : text)

  if (lang === 'bash') {
    for (const line of code.split(/(\n)/)) {
      if (line === '\n') push(null, line)
      else if (/^\s*#/.test(line)) push('tok-c', line)
      else {
        const m = /^(\s*)(\S+)(.*)$/.exec(line)
        if (m) { push(null, m[1]); push('tok-f', m[2]); push(null, m[3].replace(/\s+#.*$/, '')); const c = /\s+#.*$/.exec(m[3]); if (c) push('tok-c', c[0]) } else push(null, line)
      }
    }
    return out
  }
  if (lang === 'http') {
    for (const line of code.split(/(\n)/)) {
      const m = /^([\w-]+)(:)(.*)$/.exec(line)
      if (m) { push('tok-a', m[1]); push('tok-p', m[2]); push('tok-s', m[3]) } else push(null, line)
    }
    return out
  }

  const re = lang === 'css'
    ? /(\/\*[\s\S]*?\*\/)|("[^"]*"|'[^']*')|(--[\w-]+)|(#[0-9a-fA-F]{3,8}\b)|(\b\d+(?:\.\d+)?(?:px|%|em|rem|ms|s|vh|vw)?\b)|([{}();:,])/g
    : /(\/\/[^\n]*|\/\*[\s\S]*?\*\/|\{\/\*[\s\S]*?\*\/\})|(`(?:\\.|[^`])*`|"(?:\\.|[^"\n])*"|'(?:\\.|[^'\n])*')|(<\/?[A-Za-z][\w.]*|\/?>)|(\b\d+(?:\.\d+)?\b)|(\b[A-Za-z_$][\w$]*\b)(?=\s*\()|(\b[A-Za-z_$][\w$]*\b)|([{}()[\];,.=:?|&+*/!<>-]+)/g
  let last = 0
  for (const m of code.matchAll(re)) {
    const i = m.index ?? 0
    if (i > last) push(null, code.slice(last, i))
    const t = m[0]
    if (lang === 'css') {
      push(m[1] ? 'tok-c' : m[2] ? 'tok-s' : m[3] ? 'tok-a' : m[4] || m[5] ? 'tok-n' : 'tok-p', t)
    } else if (m[1]) push('tok-c', t)
    else if (m[2]) push('tok-s', t)
    else if (m[3]) push('tok-t', t)
    else if (m[4]) push('tok-n', t)
    else if (m[5]) push(KEYWORDS.has(t) ? 'tok-k' : 'tok-f', t)
    else if (m[6]) {
      // JSX attribute names: an identifier directly followed by `=`
      const after = code[i + t.length]
      push(KEYWORDS.has(t) ? 'tok-k' : after === '=' ? 'tok-a' : /^[A-Z]/.test(t) ? 'tok-t' : null, t)
    } else push('tok-p', t)
    last = i + t.length
  }
  if (last < code.length) push(null, code.slice(last))
  return out
}

const LABEL: Record<Lang, string> = { tsx: 'TSX', ts: 'TypeScript', css: 'CSS', bash: 'Terminal', http: 'HTTP', html: 'HTML' }

export function CopyButton({ text, className = 'copy' }: { text: string; className?: string }) {
  const [done, setDone] = useState(false)
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(text)
      setDone(true)
      setTimeout(() => setDone(false), 1600)
    } catch { /* clipboard blocked */ }
  }
  return (
    <button className={className} onClick={() => void copy()} aria-label="Copy code">
      {done ? <Icon.check /> : <Icon.copy />}{done ? 'Copied' : 'Copy'}
    </button>
  )
}

export function Code({ children, lang = 'tsx', title }: { children: string; lang?: Lang; title?: string }) {
  const code = children.replace(/^\n+|\s+$/g, '')
  return (
    <div className="code">
      <div className="code-head">
        <span>{title ?? LABEL[lang]}</span>
        <CopyButton text={code} />
      </div>
      <pre><code>{highlight(code, lang)}</code></pre>
    </div>
  )
}
