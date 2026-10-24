import { useState } from 'react'
import { Code, CopyButton } from './code'
import { Icon, VERSION, href, mount } from './layout'
import { SAMPLES, sampleUrl, type Sample } from './samples'
import { Viewer } from './viewer'
import './home.css'

const HERO_TABS: Array<[string, Sample]> = [
  ['Word', SAMPLES.find((s) => s.id === 'report')!],
  ['PowerPoint', SAMPLES.find((s) => s.id === 'launch')!],
  ['Excel', SAMPLES.find((s) => s.id === 'budget')!],
  ['Charts', SAMPLES.find((s) => s.id === 'charts')!],
]

function HeroViewer() {
  const [tab, setTab] = useState(0)
  const sample = HERO_TABS[tab][1]
  return (
    <div className="hero-viewer">
      <div className="hero-tabs" role="tablist" aria-label="Sample documents">
        {HERO_TABS.map(([label, s], i) => (
          <button key={label} role="tab" aria-selected={i === tab} className={i === tab ? 'active' : ''} onClick={() => setTab(i)}>
            <span className={`dot fmt-${s.format}`} />{label}
          </button>
        ))}
      </div>
      <Viewer key={sample.id} src={sampleUrl(sample)} height="100%" frameStyle={{ borderRadius: 0, border: 0, boxShadow: 'none' }} interactiveSheets={false} />
    </div>
  )
}

const FEATURES: Array<{ icon: keyof typeof Icon; title: string; body: string }> = [
  { icon: 'canvas', title: 'Pixel-faithful canvas rendering', body: 'Documents are parsed from their XML and painted on <canvas>, not translated into HTML. Layout, pagination and fonts stay where the author put them.' },
  { icon: 'layers', title: 'Five formats, one component', body: 'Word (.docx), Excel (.xlsx), PowerPoint (.pptx), Writer (.odt) and Rich Text (.rtf). The format is detected from the bytes.' },
  { icon: 'present', title: 'Full-screen slideshows', body: 'Present PowerPoint decks with their transitions, animations, media, Morph and 3-D effects, plus Presenter View with notes and a timer.' },
  { icon: 'chart', title: 'Native charts', body: 'Eighteen chart families, classic and Office 2016+, drawn by a lazy-loaded engine with axes, labels, trendlines and error bars.' },
  { icon: 'grid', title: 'Interactive spreadsheets', body: 'An opt-in grid viewer with tiled rendering, frozen panes, sort, filter and find, and a worker path for workbooks with millions of cells.' },
  { icon: 'search', title: 'Search, select and copy', body: 'Find with highlighted matches, text selection across pages, and copy to the clipboard, all working on the canvas.' },
  { icon: 'stream', title: 'Streaming and authenticated loading', body: 'Pass bytes, a Blob, a fetch() Response, a ReadableStream or a URL with headers, and show real download progress.' },
  { icon: 'shield', title: 'Watermarks and copy deterrents', body: 'Stamp text into every page and turn off copying and printing, from props or from server response headers.' },
  { icon: 'palette', title: 'Themeable chrome', body: 'Light and dark themes that follow the OS, with every colour exposed as a CSS custom property. The stylesheet never touches your page.' },
]

const STEPS: Array<[string, string]> = [
  ['Read', 'The file is unzipped (or, for RTF, tokenised) straight from its bytes. Nothing is uploaded.'],
  ['Parse', 'Each format is parsed into a typed model: paragraphs and runs, cells and styles, shapes and slides.'],
  ['Lay out', 'Text is measured and line-broken, tables are sized and documents paginated, the way Office does.'],
  ['Paint', 'Every page is drawn onto its own canvas, sharp at any zoom and device pixel ratio.'],
]

function Home() {
  return (
    <main>
      <section className="hero">
        <div className="container hero-grid">
          <div className="hero-copy">
            <span className="eyebrow">React component library · v{VERSION}</span>
            <h1>Office documents, rendered faithfully in the browser.</h1>
            <p className="lead">
              officeview displays Word, Excel, PowerPoint, Writer and RTF files on an HTML canvas. Pages look like they do
              in Office, on desktop and on mobile, without a server, a plugin or an HTML conversion step.
            </p>
            <div className="hero-actions">
              <a className="btn btn-primary" href={href('playground/')}>Open your own file <Icon.arrow /></a>
              <a className="btn btn-secondary" href={href('docs/')}>Read the docs</a>
            </div>
            <div className="install">
              <span className="prompt">$</span>
              <code>npm install officeview</code>
              <CopyButton text="npm install officeview" className="install-copy" />
            </div>
          </div>
          <HeroViewer />
        </div>
      </section>

      <section className="formats-strip">
        <div className="container">
          <span>Opens</span>
          <span className="fmt fmt-docx">DOCX</span>
          <span className="fmt fmt-xlsx">XLSX</span>
          <span className="fmt fmt-pptx">PPTX</span>
          <span className="fmt fmt-odt">ODT</span>
          <span className="fmt fmt-rtf">RTF</span>
          <span className="sep" />
          <span>React 18 and 19 · TypeScript types included · MIT licensed</span>
        </div>
      </section>

      <section className="section">
        <div className="container">
          <div className="section-head">
            <span className="eyebrow">Features</span>
            <h2>Everything a document viewer needs</h2>
            <p className="lead">One component covers viewing, searching, presenting and protecting documents. Heavy parts (charts, the slideshow, the spreadsheet grid) load only when a document needs them.</p>
          </div>
          <div className="feature-grid">
            {FEATURES.map((f) => {
              const I = Icon[f.icon]
              return (
                <div className="feature" key={f.title}>
                  <div className="feature-icon"><I /></div>
                  <h3>{f.title}</h3>
                  <p>{f.body}</p>
                </div>
              )
            })}
          </div>
        </div>
      </section>

      <section className="section section-alt">
        <div className="container split">
          <div>
            <span className="eyebrow">How it works</span>
            <h2 className="h2">Drawn, not converted</h2>
            <p className="lead">
              Converting a document to HTML hands its layout to the browser, which reflows text, drops page breaks and
              substitutes styles. officeview keeps the document's own geometry and paints it.
            </p>
            <ol className="pipeline">
              {STEPS.map(([title, body], i) => (
                <li key={title}>
                  <span className="num">{i + 1}</span>
                  <div><strong>{title}</strong><p>{body}</p></div>
                </li>
              ))}
            </ol>
          </div>
          <div className="quickstart">
            <h3>Quick start</h3>
            <Code lang="bash">{`npm install officeview`}</Code>
            <Code title="App.tsx">{`import { OfficeFile } from 'officeview'
import 'officeview/style.css'

export function App({ file }: { file: File }) {
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <OfficeFile data={file} />
    </div>
  )
}`}</Code>
            <p className="note">Give the viewer a sized container: it scrolls its own pages. See <a href={href('docs/#getting-started')}>Getting started</a>.</p>
          </div>
        </div>
      </section>

      <section className="section">
        <div className="container">
          <div className="section-head">
            <span className="eyebrow">Demos</span>
            <h2>See it with real documents</h2>
            <p className="lead">Each demo opens a sample file in the live viewer, with the code that produces it.</p>
          </div>
          <div className="demo-cards">
            {SAMPLES.map((s) => (
              <a className="demo-card" key={s.id} href={href(`demos/#${s.id}`)}>
                <span className={`fmt fmt-${s.format}`}>{s.format.toUpperCase()}</span>
                <h3>{s.title}</h3>
                <p>{s.summary.split('. ')[0]}.</p>
                <span className="more">Open demo <Icon.arrow /></span>
              </a>
            ))}
          </div>
        </div>
      </section>

      <section className="cta">
        <div className="container cta-inner">
          <div>
            <h2>Try it with your own file</h2>
            <p>The playground opens any .docx, .xlsx, .pptx, .odt or .rtf in your browser. Files never leave your device.</p>
          </div>
          <a className="btn btn-primary" href={href('playground/')}>Open the playground <Icon.arrow /></a>
        </div>
      </section>
    </main>
  )
}

mount('home', <Home />)
