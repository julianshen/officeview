import { useEffect, useState, type ComponentType } from 'react'
import { Code } from './code'
import { Icon, REPO, href, mount } from './layout'
import FilePicker from './examples/FilePicker'
import FilePickerSrc from './examples/FilePicker.tsx?raw'
import UrlWithProgress from './examples/UrlWithProgress'
import UrlWithProgressSrc from './examples/UrlWithProgress.tsx?raw'
import PresentControls from './examples/PresentControls'
import PresentControlsSrc from './examples/PresentControls.tsx?raw'
import PageThumbnails from './examples/PageThumbnails'
import PageThumbnailsSrc from './examples/PageThumbnails.tsx?raw'
import TextSearch from './examples/TextSearch'
import TextSearchSrc from './examples/TextSearch.tsx?raw'
import CustomStates from './examples/CustomStates'
import CustomStatesSrc from './examples/CustomStates.tsx?raw'
import BrandTheme from './examples/BrandTheme'
import BrandThemeSrc from './examples/BrandTheme.tsx?raw'
import UserWatermark from './examples/UserWatermark'
import UserWatermarkSrc from './examples/UserWatermark.tsx?raw'
import './demos.css'
import './examples.css'

interface Example {
  id: string
  title: string
  summary: string
  uses: string[]
  Component: ComponentType<{ url: string }>
  source: string
  file?: string
  name: string
}

const sample = (file: string) => href(`samples/${file}`)

const EXAMPLES: Example[] = [
  { id: 'file-picker', title: 'Open a local file', summary: 'A file input plus drag and drop. The file is read in the browser and never uploaded.', uses: ['<OfficeFile>', 'File'], Component: FilePicker, source: FilePickerSrc, name: 'FilePicker.tsx' },
  { id: 'url-progress', title: 'Download with progress', summary: 'Load a document from a URL with a custom progress bar and error message.', uses: ['<OfficeFile>', 'loading', 'error'], Component: UrlWithProgress, source: UrlWithProgressSrc, name: 'UrlWithProgress.tsx', file: 'sales-ledger.xlsx' },
  { id: 'present', title: 'Your own Present button', summary: 'Hide the built-in toolbar, start a slideshow from your UI with the ref, and track the show with callbacks.', uses: ['OfficeViewerHandle', 'present()', 'onSlideShowSlideChange'], Component: PresentControls, source: PresentControlsSrc, name: 'PresentControls.tsx', file: 'product-launch.pptx' },
  { id: 'thumbnails', title: 'Page thumbnails', summary: 'Render pages to your own canvases without the viewer, for previews, galleries or a page picker.', uses: ['loadOfficeFile()', 'getPaintables()'], Component: PageThumbnails, source: PageThumbnailsSrc, name: 'PageThumbnails.tsx', file: 'quarterly-review.docx' },
  { id: 'text-search', title: 'Search without displaying', summary: 'Build the viewer\'s text index headlessly and list matches with their context and page.', uses: ['buildTextIndex()', 'findMatches()'], Component: TextSearch, source: TextSearchSrc, name: 'TextSearch.tsx', file: 'quarterly-review.docx' },
  { id: 'custom-states', title: 'Custom loading and errors', summary: 'Own the loading, error and notice UI with the useOfficeFile hook, then render with <OfficeDoc>.', uses: ['useOfficeFile()', '<OfficeDoc>'], Component: CustomStates, source: CustomStatesSrc, name: 'CustomStates.tsx', file: 'operating-budget.xlsx' },
  { id: 'brand-theme', title: 'Brand colours', summary: 'Re-brand the viewer chrome in light and dark themes by overriding its CSS tokens.', uses: ['theme', '--ov-color-*'], Component: BrandTheme, source: BrandThemeSrc, name: 'BrandTheme.tsx', file: 'product-launch.pptx' },
  { id: 'user-watermark', title: 'Per-user watermark', summary: 'Stamp each page with the viewer\'s identity and the date, and turn off copy and print.', uses: ['watermark', 'allowCopy', 'allowPrint'], Component: UserWatermark, source: UserWatermarkSrc, name: 'UserWatermark.tsx', file: 'quarterly-review.docx' },
]

const fromHash = () => EXAMPLES.find((e) => `#${e.id}` === location.hash) ?? EXAMPLES[0]

function ExampleView({ example }: { example: Example }) {
  const { Component } = example
  return (
    <article className="demo">
      <header className="demo-head">
        <div className="demo-meta">{example.uses.map((u) => <span key={u} className="badge accent">{u}</span>)}</div>
        <h1>{example.title}</h1>
        <p className="lead">{example.summary}</p>
      </header>
      <div className="example-frame" key={example.id}>
        <Component url={example.file ? sample(example.file) : ''} />
      </div>
      <h2 className="code-title">Source</h2>
      <Code title={example.name}>{example.source}</Code>
      <p className="demo-foot">
        Remember to import <code>officeview/style.css</code> once in your app. Full projects for Vite and Next.js are in{' '}
        <a href={`${REPO}/tree/main/examples`}>examples/</a>.
      </p>
    </article>
  )
}

function Examples() {
  const [example, setExample] = useState(fromHash)
  useEffect(() => {
    const on = () => setExample(fromHash())
    window.addEventListener('hashchange', on)
    return () => window.removeEventListener('hashchange', on)
  }, [])
  useEffect(() => { document.title = `${example.title} — officeview examples` }, [example])
  return (
    <main className="container demos-layout">
      <nav className="demo-nav" aria-label="Examples">
        <div className="demo-nav-title">Examples</div>
        {EXAMPLES.map((e) => (
          <a key={e.id} href={`#${e.id}`} aria-current={e.id === example.id ? 'page' : undefined}>
            <span>{e.title}</span>
          </a>
        ))}
        <div className="demo-nav-title nav-sub">Projects</div>
        <a href={`${REPO}/tree/main/examples/vite-react`}><Icon.github /> Vite + React</a>
        <a href={`${REPO}/tree/main/examples/nextjs`}><Icon.github /> Next.js</a>
      </nav>
      <ExampleView example={example} />
    </main>
  )
}

mount('examples', <Examples />)
