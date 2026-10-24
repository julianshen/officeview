import { useEffect, useRef, useState, type RefObject } from 'react'
import { SlideShowError, type OfficeViewerHandle } from 'officeview'
import { Code } from './code'
import { Icon, href, mount } from './layout'
import { FORMAT_LABEL, SAMPLES, sampleUrl, type Sample } from './samples'
import { Viewer } from './viewer'
import './demos.css'

const fromHash = () => SAMPLES.find((s) => `#${s.id}` === location.hash) ?? SAMPLES[0]

function PresentButton({ viewer }: { viewer: RefObject<OfficeViewerHandle | null> }) {
  const [message, setMessage] = useState('')
  return (
    <span className="present">
      <button
        className="btn btn-primary btn-sm"
        onClick={() => {
          setMessage('')
          // present() must be called synchronously inside the click handler.
          viewer.current?.present({ from: 'beginning' }).catch((e) => {
            if (e instanceof SlideShowError && e.code !== 'cancelled') setMessage(`Could not start the show (${e.code}).`)
          })
        }}
      >
        <Icon.play /> Present from the beginning
      </button>
      {message && <span className="present-msg">{message}</span>}
    </span>
  )
}

function Demo({ sample }: { sample: Sample }) {
  const viewer = useRef<OfficeViewerHandle>(null)
  return (
    <article className="demo" aria-labelledby="demo-title">
      <header className="demo-head">
        <div className="demo-meta">
          <span className={`fmt fmt-${sample.format}`}>{sample.format.toUpperCase()}</span>
          <span className="badge">{FORMAT_LABEL[sample.format]}</span>
          {sample.props?.interactiveSheets && <span className="badge accent">interactiveSheets</span>}
          {sample.props?.watermark && <span className="badge accent">watermark</span>}
        </div>
        <h1 id="demo-title">{sample.title}</h1>
        <p className="lead">{sample.summary}</p>
      </header>
      <div className="demo-toolbar">
        <ul className="try">
          {sample.tryThis.map((t) => <li key={t}>{t}</li>)}
        </ul>
        <div className="demo-actions">
          {sample.format === 'pptx' && <PresentButton viewer={viewer} />}
          <a className="btn btn-secondary btn-sm" href={sampleUrl(sample)} download><Icon.download /> {sample.file}</a>
        </div>
      </div>
      <Viewer key={sample.id} ref={viewer} src={sampleUrl(sample)} height="min(760px, calc(100vh - 140px))" {...sample.props} />
      <h2 className="code-title">Code</h2>
      <Code title="Demo.tsx">{`import { OfficeFile${sample.format === 'pptx' ? ', type OfficeViewerHandle' : ''} } from 'officeview'
import 'officeview/style.css'

${sample.code}`}</Code>
      <p className="demo-foot">
        Want to change the options? Open this file in the <a href={href(`playground/#${sample.id}`)}>playground</a>.
      </p>
    </article>
  )
}

function Demos() {
  const [sample, setSample] = useState(fromHash)
  useEffect(() => {
    const on = () => setSample(fromHash())
    window.addEventListener('hashchange', on)
    return () => window.removeEventListener('hashchange', on)
  }, [])
  useEffect(() => {
    document.title = `${sample.title} — officeview demos`
  }, [sample])
  return (
    <main className="container demos-layout">
      <nav className="demo-nav" aria-label="Demos">
        <div className="demo-nav-title">Demos</div>
        {SAMPLES.map((s) => (
          <a key={s.id} href={`#${s.id}`} aria-current={s.id === sample.id ? 'page' : undefined}>
            <span className={`dot fmt-${s.format}`} />
            <span>{s.title}</span>
          </a>
        ))}
      </nav>
      <Demo sample={sample} />
    </main>
  )
}

mount('demos', <Demos />)
