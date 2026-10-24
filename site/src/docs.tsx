import { Code } from './code'
import { DocShell, H3, Note, Props, Section, Table, c, type DocNav, type PropRow } from './doc-kit'
import { REPO, href, mount } from './layout'

const NAV: DocNav = [
  ['Getting started', [['introduction', 'Introduction'], ['installation', 'Installation'], ['getting-started', 'Quick start'], ['sizing', 'Sizing the viewer'], ['formats', 'Supported formats']]],
  ['Components', [['officefile', '<OfficeFile>'], ['officedoc', '<OfficeDoc>'], ['props', 'Props reference'], ['viewer-handle', 'OfficeViewerHandle']]],
  ['Loading documents', [['sources', 'Byte sources'], ['progress', 'Progress'], ['authentication', 'Authenticated downloads'], ['load-office-file', 'loadOfficeFile()'], ['use-office-file', 'useOfficeFile()']]],
  ['Features', [['theming', 'Theming'], ['search', 'Search and selection'], ['protection', 'Copy and print'], ['watermarks', 'Watermarks'], ['server-headers', 'Server headers'], ['slideshow', 'Slideshow'], ['spreadsheets', 'Interactive spreadsheets'], ['charts', 'Charts'], ['fonts', 'Fonts']]],
  ['Reference', [['low-level', 'Low-level API'], ['types', 'Types'], ['limitations', 'Known limitations']]],
]

// ---------------------------------------------------------------------------
// Content
// ---------------------------------------------------------------------------

const VIEW_PROPS: PropRow[] = [
  ['theme', "'light' | 'dark' | 'system'", "'system'", <>Colour theme of the viewer chrome. {c("'system'")} follows {c('prefers-color-scheme')}. The document itself is never themed. See <a href="#theming">Theming</a>.</>],
  ['title', 'string', '', <>File name shown in the toolbar. {c('<OfficeFile>')} derives one from a {c('File')}, the URL or {c('Content-Disposition')}.</>],
  ['showToolbar', 'boolean', 'true', 'Show the toolbar and status line of the page viewer (Word, PowerPoint, RTF, ODT and page-mode Excel).'],
  ['showZoomControls', 'boolean', 'true', 'Show the zoom controls.'],
  ['showSearch', 'boolean', 'true', <>Show Find. {c('Ctrl/⌘+F')} opens it while the viewer has focus.</>],
  ['pageGapPx', 'number', '', 'Gap between pages, in CSS pixels.'],
  ['background', 'string', '', <>Backdrop colour around the pages. Overrides the {c('--ov-color-backdrop')} token.</>],
  ['className', 'string', '', 'Class name of the scrolling page region.'],
  ['style', 'CSSProperties', '', 'Inline style of the scrolling page region.'],
  ['allowCopy', 'boolean', 'true', <>{c('false')} removes text selection, the Copy button and the copy shortcut. See <a href="#protection">Copy and print</a>.</>],
  ['allowPrint', 'boolean', 'true', <>{c('false')} hides the Print button and removes the viewer from the browser's print output.</>],
  ['watermark', 'WatermarkOptions', '', <>Text painted into every page. See <a href="#watermarks">Watermarks</a>.</>],
  ['fallbackFonts', 'FallbackFontsOptions', '', <>Pin explicit font fallback chains. See <a href="#fonts">Fonts</a>.</>],
  ['interactiveSheets', 'boolean', 'false', <>Open Excel files in the interactive grid viewer (experimental). See <a href="#spreadsheets">Interactive spreadsheets</a>.</>],
  ['allowPresent', 'boolean', 'true', <>Allow presenting PowerPoint decks. {c('false')} hides Present and makes {c('present()')} reject.</>],
  ['presentShortcuts', 'boolean', 'true', <>{c('F5')} / {c('Shift+F5')} start a show while focus is inside the viewer.</>],
  ['slideShowMotion', "'auto' | 'full' | 'reduced'", "'auto'", <>{c("'auto'")} follows {c('prefers-reduced-motion')}; {c("'reduced'")} turns transitions into short crossfades.</>],
  ['onSlideShowStart', '() => void', '', 'A show has started.'],
  ['onSlideShowEnd', '() => void', '', <>Fires exactly once per started show, before {c('session.ended')} resolves.</>],
  ['onSlideShowSlideChange', '(slideIndex: number, buildIndex: number) => void', '', 'The slide or the click checkpoint changed.'],
  ['onSlideShowDiagnostic', '(d: SlideShowDiagnostic) => void', '', 'An effect was approximated, replaced or refused. Deduplicated per effect and slide.'],
]

const TOKENS: Array<[string, string, string, string]> = [
  ['surface', '#ffffff', '#292827', 'Menus, dialogs, popovers'],
  ['chrome', '#faf9f8', '#201f1e', 'Toolbar and status line'],
  ['field', '#ffffff', '#1b1a19', 'Text inputs'],
  ['border', '#e1dfdd', '#3b3a39', 'Control and panel borders'],
  ['text', '#201f1e', '#f3f2f1', 'Primary chrome text'],
  ['muted', '#605e5c', '#a19f9d', 'Secondary text'],
  ['icon', '#323130', '#d2d0ce', 'Toolbar icons'],
  ['hover', '#f3f2f1', '#3b3a39', 'Hover and pressed background'],
  ['accent', '#0f6cbd', '#479ef5', 'Focus ring, selected controls, primary buttons'],
  ['on-accent', '#ffffff', '#1b1a19', 'Text on accent'],
  ['header', '#f3f2f1', '#201f1e', 'Spreadsheet header background'],
  ['header-text', '#323130', '#d2d0ce', 'Spreadsheet header labels'],
  ['backdrop', '#f0efee', '#161514', 'Area around the pages'],
  ['selection', '#107c41', '#107c41', 'Cell selection border (on paper)'],
  ['find', 'rgba(255,214,0,.45)', 'same', 'Find match highlight (on paper)'],
  ['find-active', '#ca5010', '#ca5010', 'Active match outline (on paper)'],
  ['text-selection', 'rgba(64,128,255,.35)', 'same', 'Text selection (on paper)'],
]

function Swatch({ value }: { value: string }) {
  return value === 'same' ? <span className="muted">same</span> : <span className="swatch"><span style={{ background: value }} /><code>{value}</code></span>
}

function Content() {
  return (
    <>
      <Section id="introduction" title="Introduction">
        <p className="doc-lead">
          officeview is a set of React components that render Office documents on an HTML {c('<canvas>')}. It parses
          the file in the browser, lays it out the way Office does, and paints each page. Nothing is uploaded and no
          HTML conversion takes place, so pages keep their geometry, pagination and fonts.
        </p>
        <p>
          Most applications need one component, {c('<OfficeFile>')}, which loads a file from almost any source and shows
          it with a toolbar, zoom, find, selection, printing and, for PowerPoint, a full-screen slideshow. Lower-level
          functions are exported for hosts that want to parse or paint documents themselves.
        </p>
        <div className="doc-cards">
          <a href={href('guides/')}><strong>Guides</strong><span>Vite, Next.js, workers, serving files, CSP</span></a>
          <a href={href('examples/')}><strong>Examples</strong><span>Runnable recipes with their source</span></a>
          <a href={href('playground/')}><strong>Playground</strong><span>Open your own file, try every option</span></a>
        </div>
      </Section>

      <Section id="installation" title="Installation">
        <Code lang="bash">{`npm install officeview
# or
pnpm add officeview
yarn add officeview
bun add officeview`}</Code>
        <p>
          React 18 or later is a peer dependency. The package is ESM-only and ships TypeScript declarations. The
          toolbar uses {c('lucide-react')} and {c('motion')} (pinned to {c('14.4.0')}), both loaded in lazy chunks.
        </p>
      </Section>

      <Section id="getting-started" title="Quick start">
        <p>Import the component and, once per application, its stylesheet:</p>
        <Code title="App.tsx">{`import { useState } from 'react'
import { OfficeFile } from 'officeview'
import 'officeview/style.css'

export function App() {
  const [file, setFile] = useState<File | null>(null)
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <input type="file" accept=".docx,.xlsx,.pptx,.odt,.rtf" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
      <OfficeFile data={file} />
    </div>
  )
}`}</Code>
        <Note>
          <p>
            <strong>The stylesheet is required.</strong> Without {c('officeview/style.css')} the viewer chrome is unstyled and
            the console shows {c('[officeview] style.css not loaded')}. The stylesheet has no reset and every rule is scoped
            to the viewer, so it never restyles your page.
          </p>
        </Note>
        <p>To load a file from a URL, pass a descriptor. The download streams, and the viewer shows progress. In a component that re-renders, create the descriptor once (see <a href="#sources">Byte sources</a>):</p>
        <Code>{`<OfficeFile data={{ url: '/files/report.docx' }} />`}</Code>
      </Section>

      <Section id="sizing" title="Sizing the viewer">
        <p>
          The viewer scrolls its own container: page tracking (“3 / 12”) and Previous / Next move that container, never
          the host page. Give it a constrained height, for example as a flex item in a sized column:
        </p>
        <Code lang="css">{`.viewer-host {
  height: 100vh;            /* or any fixed height */
  display: flex;
  flex-direction: column;
}
.viewer-host > [data-ov-theme-root] {
  flex: 1 1 auto;
  min-height: 0;
}`}</Code>
        <p>
          In a host that lets the viewer grow to its full content height, the container never scrolls, so the page
          readout stays on the first page. The interactive spreadsheet viewer defaults to 400 px tall in an unsized host.
        </p>
      </Section>

      <Section id="formats" title="Supported formats">
        <p>The format is detected from the file's bytes, not its name. All five go through the same component.</p>
        <Table
          head={['Format', 'Rendered']}
          rows={[
            [<span className="fmt fmt-docx">DOCX</span>, 'Paragraph and character styles, lists, tables (spans, merges, repeating headers, shading, borders), inline and floating images, headers and footers with PAGE/NUMPAGES fields, sections, charts.'],
            [<span className="fmt fmt-xlsx">XLSX</span>, 'Sheets, shared strings, number formats, fonts, fills, borders, merged cells, column widths and row heights, formulas (built-in engine), images, charts. Optional interactive grid viewer.'],
            [<span className="fmt fmt-pptx">PPTX</span>, 'Shapes, text bodies, pictures with cropping, tables with table styles, master and layout backgrounds and artwork, charts. Full-screen slideshow with transitions, animations and media.'],
            [<span className="fmt fmt-odt">ODT</span>, 'Paragraph and text styles, headings and outline numbering, lists, tables, images, page layouts, headers and footers, page fields. Laid out by the Word engine.'],
            [<span className="fmt fmt-rtf">RTF</span>, 'Runs, font and colour tables, alignment, spacing, sections, Unicode and code-page escapes, tables, PNG and JPEG pictures, list paragraphs.'],
          ]}
        />
      </Section>

      <Section id="officefile" title="<OfficeFile>" lead="Loads a document from a byte source and renders it. This is the component most applications use.">
        <Code>{`import { OfficeFile } from 'officeview'

<OfficeFile
  data={source}                       // see Byte sources
  loading={({ loaded, total }) => <Progress loaded={loaded} total={total} />}
  error={(message) => <p role="alert">{message}</p>}
  theme="system"
/>`}</Code>
        <p>
          {c('<OfficeFile>')} accepts every <a href="#props">viewer prop</a> plus the ones below. It forwards a{' '}
          <a href="#viewer-handle">ref</a> to the viewer.
        </p>
        <Props rows={[
          ['data', 'ByteSource | null | undefined', '', <>The file. See <a href="#sources">Byte sources</a>. Changing it loads the new file. While it is {c('null')}, the loading state is shown.</>],
          ['loading', 'ReactElement | ((p: Progress) => ReactElement) | null', '', 'Shown while loading. The function form receives live download progress.'],
          ['error', '(message: string) => ReactElement', '', <>Shown when the file cannot be read. Use {c('useOfficeFile')} for structured error codes.</>],
          ['viewerWorker', '() => Worker', '', <>Factory for the module worker that opens large spreadsheets with {c('interactiveSheets')}. See <a href="#spreadsheets">Interactive spreadsheets</a>.</>],
        ]} />
      </Section>

      <Section id="officedoc" title="<OfficeDoc>" lead="Renders a document that has already been parsed. Use it when you load files yourself.">
        <Code>{`import { OfficeDoc, loadOfficeFile } from 'officeview'

const doc = await loadOfficeFile(bytes)

<OfficeDoc document={doc} theme="dark" />`}</Code>
        <Props rows={[
          ['document', 'DocxDocument | XlsxDocument | PptxDocument', '', <>A parsed document, from {c('loadOfficeFile')} or a {c('parse*')} function. ODT and RTF parse to {c('DocxDocument')}.</>],
        ]} />
        <p>
          {c('<OfficeDoc>')} accepts every <a href="#props">viewer prop</a>. It has no {c('viewerWorker')}: large
          spreadsheets should be parsed from bytes through {c('<OfficeFile>')} or {c('useOfficeFile')}.
        </p>
      </Section>

      <Section id="props" title="Props reference" lead="Shared by <OfficeFile> and <OfficeDoc>. All are optional.">
        <Props rows={VIEW_PROPS} />
      </Section>

      <Section id="viewer-handle" title="OfficeViewerHandle" lead="Both components forward a ref with one method, present(), to start a slideshow from your own UI.">
        <Code>{`import { useRef } from 'react'
import { OfficeFile, SlideShowError, type OfficeViewerHandle } from 'officeview'

const viewer = useRef<OfficeViewerHandle>(null)

<button onClick={() => {
  // Call present() synchronously in the click handler: no await before it.
  viewer.current?.present({ from: 'beginning' })
    .then((session) => session.ended)
    .catch((e) => { if (e instanceof SlideShowError) console.warn(e.code) })
}}>Present</button>
<OfficeFile ref={viewer} data={deck} />`}</Code>
        <p>See <a href="#slideshow">Slideshow</a> for {c('present()')} options, the session object and error codes.</p>
      </Section>

      <Section id="sources" title="Byte sources" lead="data accepts anything byte-shaped, so you can hand the viewer a download instead of waiting for it.">
        <Code>{`<OfficeFile data={bytes} />                         // ArrayBuffer | Uint8Array
<OfficeFile data={file} />                          // Blob or File
<OfficeFile data={fetch('/deck.pptx')} />           // a Response, or a promise of one
<OfficeFile data={response.body!} />                // ReadableStream<Uint8Array>
<OfficeFile data={{ url: '/report.docx' }} />       // a URL descriptor
<OfficeFile data={new Request('/report.docx')} />   // a Request`}</Code>
        <Table
          head={['Source', 'Notes']}
          rows={[
            [c('ArrayBuffer | Uint8Array'), 'Bytes already in memory.'],
            [c('Blob'), <>Including {c('File')}. Its size gives a progress total.</>],
            [c('Response'), <>Non-2xx statuses fail fast with {c('HTTP <status> while downloading the document')}.</>],
            [c('ReadableStream<Uint8Array>'), <>Progress reports {c('loaded')} with no {c('total')}.</>],
            [c('{ url, headers?, credentials? }'), <>Downloaded with {c('fetch')}, streaming progress, optional headers and credentials.</>],
            [c('Request'), 'Carries method, mode and cache settings with it.'],
            [c('Promise<any of these>'), <>For example {c('fetch(url)')} itself.</>],
          ]}
        />
        <Note warn>
          <p>
            <strong>Keep the source object stable.</strong> {c('data')} is compared by identity: a new object loads the file
            again. An inline {c("data={{ url }}")} in a component that re-renders re-downloads the document each time. Create
            the descriptor once, with {c('useMemo')} or outside the component:
          </p>
          <Code>{`const source = useMemo(() => ({ url, headers: { Authorization: \`Bearer \${token}\` } }), [url, token])

<OfficeFile data={source} />`}</Code>
        </Note>
        <Note>
          <p>
            Rendering starts once the last byte has arrived: DOCX, XLSX and PPTX are zip files whose index sits at the end.
            Streaming makes the download visible, not the first page faster.
          </p>
        </Note>
        <p>
          Bytes, Blobs and URL descriptors can be read again, so changing {c('interactiveSheets')} or{' '}
          {c('viewerWorker')} reopens them. A {c('Response')}, stream, {c('Request')} or promise is consumed by its first
          load; pass a fresh one to reload.
        </p>
      </Section>

      <Section id="progress" title="Progress">
        <p>Pass a function as {c('loading')} to receive {c('{ loaded, total }')} as bytes arrive:</p>
        <Code>{`<OfficeFile
  data={{ url: '/files/large.xlsx' }}
  loading={({ loaded, total }) =>
    total ? <Progress value={loaded / total} /> : <Spinner label={\`\${loaded} bytes\`} />
  }
/>`}</Code>
        <p>{c('total')} is present when the source declares a size: a {c('Blob')}, or a response with {c('Content-Length')}.</p>
      </Section>

      <Section id="authentication" title="Authenticated downloads">
        <Code>{`<OfficeFile data={{ url: '/api/files/42', headers: { Authorization: \`Bearer \${token}\` } }} />

// cookies on a cross-site request
<OfficeFile data={{ url: 'https://files.example.com/report.docx', credentials: 'include' }} />

// or do the fetch yourself
<OfficeFile data={fetch(url, { headers: { Authorization: \`Bearer \${token}\` } })} />`}</Code>
        <p>Request headers are never copied into error messages.</p>
      </Section>

      <Section id="load-office-file" title="loadOfficeFile()" lead="Load and parse a file without rendering it.">
        <Code lang="ts">{`function loadOfficeFile(
  source: ByteSource,
  options?: {
    onProgress?: (p: Progress) => void
    onProtection?: (policy: ProtectionPolicy) => void
    onWatermark?: (mark: WatermarkOptions | undefined) => void
  },
): Promise<DocxDocument | XlsxDocument | PptxDocument>`}</Code>
        <Code lang="ts">{`import { loadOfficeFile } from 'officeview'

const doc = await loadOfficeFile(fetch('/report.docx'), {
  onProgress: ({ loaded, total }) => console.log(loaded, total),
})`}</Code>
        <p>
          The format is detected from the bytes. A spreadsheet that exceeds the in-memory limits rejects with a typed
          refusal; use {c('<OfficeFile>')} with a worker for those.
        </p>
      </Section>

      <Section id="use-office-file" title="useOfficeFile()" lead="The hook behind <OfficeFile>, for hosts that render their own loading and error states.">
        <Code>{`import { useMemo } from 'react'
import { OfficeDoc, useOfficeFile } from 'officeview'

function Viewer({ url }: { url: string }) {
  const source = useMemo(() => ({ url }), [url])   // a stable source: see Byte sources
  const state = useOfficeFile(source)
  if (state.status === 'loading') return <Spinner />
  if (state.status === 'error') return <ErrorPanel message={state.message} code={state.code} />
  return (
    <>
      {state.notices?.map((n) => <Banner key={n}>{n}</Banner>)}
      <OfficeDoc document={state.document} />
    </>
  )
}`}</Code>
        <Table
          head={['status', 'Fields']}
          rows={[
            [c("'loading'"), '—'],
            [c("'ready'"), <>{c('document')}, and {c('notices')}: messages such as “Showing saved values”, which you should display.</>],
            [c("'error'"), <>{c('message')}, and for refused spreadsheets {c('code')} and {c('reason')}.</>],
          ]}
        />
        <p>
          Options: {c('onProgress')}, {c('onProtection')}, {c('onWatermark')}, {c('onHeaders')} (the response headers of
          the download), {c('interactiveSheets')} and {c('viewerWorker')}.
        </p>
      </Section>

      <Section id="theming" title="Theming">
        <p>
          The viewer has a light and a dark theme and follows the operating system by default. Set {c('theme')} to force one.
          Only the chrome is themed: pages, sheets and charts stay white paper in both themes.
        </p>
        <p>
          Every colour is a CSS custom property named {c('--ov-color-<name>')}, defined with zero specificity on the
          viewer's theme root ({c('[data-ov-theme-root]')}). Any rule that targets the root overrides it, without{' '}
          {c('!important')}:
        </p>
        <Code lang="css">{`.brand [data-ov-theme-root] {
  --ov-color-accent: #c50f1f;
}
.brand [data-ov-theme-root][data-ov-theme="dark"] {
  --ov-color-accent: #ff99a4;
}`}</Code>
        <Table
          head={['Token', 'Light', 'Dark', 'Used for']}
          rows={TOKENS.map(([name, light, dark, use]) => [c(`--ov-color-${name}`), <Swatch value={light} />, <Swatch value={dark} />, use])}
        />
        <p>
          Radii and shadows: {c('--ov-radius-control')} (4px), {c('--ov-radius-panel')} (6px), {c('--ov-shadow-rest')}{' '}
          and {c('--ov-shadow-popover')}. The defaults meet WCAG contrast in both themes; check your own values against the same
          rules (text 4.5:1, controls 3:1).
        </p>
        <Note>
          <p>
            For a {c("theme=\"system\"")} viewer, put dark overrides inside {c('@media (prefers-color-scheme: dark)')} and
            target {c('[data-ov-theme-root][data-ov-theme=system]')}.
          </p>
        </Note>
      </Section>

      <Section id="search" title="Search and selection">
        <p>
          Press {c('Ctrl/⌘+F')} in the viewer, or use the toolbar, to search. {c('Enter')} and {c('Shift+Enter')} step through
          matches; {c('Esc')} clears. Matches are highlighted on the canvas and the active one is scrolled into view.
        </p>
        <p>
          Drag to select text, double-click for a word and triple-click for a line. A selection can span pages. Copy with{' '}
          {c('Ctrl/⌘+C')} or the toolbar's Copy button. Search and selection work the same way in every format.
        </p>
      </Section>

      <Section id="protection" title="Copy and print">
        <Code>{`<OfficeFile data={source} allowCopy={false} allowPrint={false} />`}</Code>
        <ul>
          <li>{c('allowCopy={false}')} removes text selection, the copy shortcut and the Copy button, and blocks native copy and the context menu inside the viewer.</li>
          <li>{c('allowPrint={false}')} hides the Print button, blocks {c('Ctrl/⌘+P')} while the viewer has focus, and removes the viewer from the browser's print output. Other content on the page still prints.</li>
        </ul>
        <p>
          When printing is allowed, Print renders each page to an image and prints only those images, so your page's
          content never reaches the printout.
        </p>
        <Note warn>
          <p>
            These are deterrents, not DRM. A reader can still take screenshots or use developer tools. If a document is
            confidential, serve it only to authorised users and treat the viewer as a preview.
          </p>
        </Note>
      </Section>

      <Section id="watermarks" title="Watermarks">
        <Code>{`<OfficeFile
  data={source}
  watermark={{ text: 'CONFIDENTIAL', placement: 'tile', rotate: -45, opacity: 0.15 }}
/>`}</Code>
        <Props rows={[
          ['text', 'string', '', 'The mark. Empty text disables the watermark.'],
          ['placement', "'tile' | 'center' | 'header' | 'footer'", "'tile'", 'Where the mark is painted on each page.'],
          ['rotate', 'number', '-45', 'Degrees clockwise.'],
          ['opacity', 'number', '0.15', <>Clamped to {c('0..1')}.</>],
          ['color', 'string', '', 'Any CSS colour. When unset, a colour is chosen from the page background so the mark stays visible.'],
          ['fontSizePt', 'number', '', '7% of the page width when unset.'],
          ['fontFamily', 'string', '', 'Font family of the mark.'],
        ]} />
        <p>
          The mark is painted into the page bitmap underneath the content, so hiding an element does not remove it. It is
          excluded from search.
        </p>
      </Section>

      <Section id="server-headers" title="Server headers" lead="When the viewer downloads the file itself, the server can impose protection and a watermark.">
        <Code lang="http">{`X-OfficeView-Protection: no-copy, no-print
X-OfficeView-Watermark: text=CONFIDENTIAL; placement=tile; rotate=-45; opacity=0.15`}</Code>
        <ul>
          <li>{c('X-OfficeView-Protection')} tokens deny one capability each. The header and the props merge most-restrictively: a server denial cannot be re-enabled from props.</li>
          <li>{c('X-OfficeView-Watermark')} takes {c('key=value')} pairs separated by {c(';')} ({c('text')}, {c('placement')}, {c('rotate')}, {c('opacity')}, {c('color')}, {c('fontSize')}, {c('font')}). Keys are case-insensitive and values are percent-decoded; a {c(';')} inside the text is written {c('%3B')}. The header wins over the {c('watermark')} prop.</li>
          <li>A missing or malformed header changes nothing, so existing servers need no changes.</li>
        </ul>
        <p>
          The policy is sealed from the response that delivered the document and resets when {c('data')} changes. For
          cross-origin downloads, expose the headers with {c('Access-Control-Expose-Headers')}.
        </p>
      </Section>

      <Section id="slideshow" title="Slideshow">
        <p>
          PowerPoint decks can be presented full screen, with their transitions (including 3-D transitions and Morph),
          animations, timings, triggers, audio and video, and with Presenter View. The show engine is a lazy chunk; it
          draws with WebGL2 and falls back to Canvas 2D.
        </p>
        <H3>Starting a show</H3>
        <ul>
          <li>The <strong>Present</strong> button in the toolbar, with “From beginning”, “From current slide” and “Presenter View” in its menu.</li>
          <li>{c('F5')} from the beginning, {c('Shift+F5')} from the current slide, {c('Alt+F5')} with Presenter View, while focus is inside the viewer.</li>
          <li>Your own UI, through the <a href="#viewer-handle">ref</a>.</li>
        </ul>
        <H3>present(options)</H3>
        <Props rows={[
          ['from', "'beginning' | 'current' | number", "'beginning'", 'Where to start; a number is a 0-based slide index.'],
          ['presenterView', "boolean | 'auto'", "'auto'", <>{c('true')} opens Presenter View in a second window. {c("'auto'")} does so only when a two-screen placement is ready. {c('false')} shows the audience view only.</>],
        ]} />
        <Note>
          <p>
            Browsers only allow full screen during a user gesture. Call {c('present()')} synchronously inside the click
            or key handler. If full screen is refused, the show runs in the window instead.
          </p>
        </Note>
        <H3>SlideShowSession</H3>
        <p>{c('present()')} resolves to a session:</p>
        <Code lang="ts">{`interface SlideShowSession {
  readonly slideIndex: number                    // 0-based
  readonly buildIndex: number                    // click checkpoint within the slide
  readonly surface: 'fullscreen' | 'in-window'
  readonly presenter: 'window' | 'swap' | 'none'
  next(): void
  previous(): void
  goTo(slideIndex: number): void
  exit(): void                                   // idempotent
  readonly ended: Promise<void>                  // resolves once; never rejects
}`}</Code>
        <H3>Errors</H3>
        <p>{c('present()')} rejects with a {c('SlideShowError')} whose {c('code')} is one of:</p>
        <Table
          head={['code', 'When']}
          rows={[
            [c('gesture-required'), 'The browser reports no user activation for the call.'],
            [c('not-presentable'), <>The document is not a PowerPoint deck, or {c('allowPresent')} is false.</>],
            [c('already-presenting'), 'A show is already starting or running on the page.'],
            [c('load-failed'), 'The slideshow chunk or the show data could not be loaded.'],
            [c('cancelled'), 'The start was cancelled: Esc, unmount or a document change.'],
            [c('invalid-slide'), <>{c('from')} is not a slide index in range.</>],
          ]}
        />
        <H3>Keys during a show</H3>
        <Table
          head={['Action', 'Input']}
          rows={[
            ['Next build or slide', 'Click, Space, Enter, →, ↓, Page Down, N, swipe left'],
            ['Previous', '←, ↑, Page Up, P, Backspace, swipe right'],
            ['First / last slide', 'Home / End'],
            ['Go to slide n', 'Type the number, then Enter'],
            ['Black / white screen', 'B or . / W or ,'],
            ['Exit', 'Esc or -'],
          ]}
        />
        <H3>Embedding in an iframe</H3>
        <p>
          Full screen needs {c('allow="fullscreen"')}, Presenter View's window needs {c('allow-popups')} in a sandbox, and
          two-screen placement needs {c('allow="window-management"')}. Without them the show still runs, in the window or
          on one screen, and reports a diagnostic.
        </p>
        <Note>
          <p>
            Every effect that is approximated, replaced or refused is reported through {c('onSlideShowDiagnostic')}. The
            full effect catalogue is in <a href={`${REPO}/blob/main/docs/validation/slideshow.md`}>docs/validation/slideshow.md</a>.
          </p>
        </Note>
      </Section>

      <Section id="spreadsheets" title="Interactive spreadsheets">
        <p>
          {c('interactiveSheets')} opens Excel files in a grid viewer with tiled rendering, frozen panes, 25%–400% zoom,
          selection, sorting, filtering, find and copy. It is experimental and off by default; with it off, spreadsheets render
          as fit-to-width pages and the grid code is never loaded.
        </p>
        <Code>{`// viewer-worker.ts
import 'officeview/worker'`}</Code>
        <Code>{`import { OfficeFile } from 'officeview'

const createViewerWorker = () =>
  new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })

<OfficeFile data={file} interactiveSheets viewerWorker={createViewerWorker} />`}</Code>
        <p>
          The worker is optional. Small workbooks open in memory; a workbook admitted as large opens in the worker, which
          keeps its pages and serves the grid bounded windows. Without a worker, large files open in a compatibility mode.
          With Vite, set {c("worker: { format: 'es' }")}.
        </p>
        <ul>
          <li><strong>View-only.</strong> Sorting, filtering and resizing never change the file.</li>
          <li><strong>Formulas.</strong> Small files are recalculated by the built-in engine; large files show their saved values, with a notice.</li>
          <li><strong>Capacity.</strong> The full 1,048,576 × 16,384 grid is addressable; memory budgets differ for desktop and mobile.</li>
        </ul>
        <p>The README covers admission limits, filter semantics and accessibility in detail.</p>
      </Section>

      <Section id="charts" title="Charts">
        <p>
          Charts in Word, Excel and PowerPoint files are drawn by one engine, loaded the first time a document contains a
          chart. It reads the chart's saved values, the worksheet or the embedded workbook. External workbook links are
          never fetched.
        </p>
        <Table
          head={['Kind', 'Families']}
          rows={[
            ['Classic', 'bar and column, line, area, pie, doughnut, pie of pie, scatter, bubble, radar, contour, stock (3-D variants are drawn flat)'],
            ['Office 2016+', 'waterfall, funnel, treemap, sunburst, histogram, Pareto, box and whisker'],
            ['Shown another way', '3-D surface (a card); region maps (the file\'s fallback picture)'],
          ]}
        />
        <p>
          Axes (category, value, date, logarithmic, secondary), data labels, markers, trendlines, error bars, titles and
          legends are supported. A chart the engine cannot draw keeps the file's fallback picture, or shows a card; it is
          never silently blank.
        </p>
      </Section>

      <Section id="fonts" title="Fonts">
        <p>
          Text is measured and drawn with the fonts available to the browser. When a document names a font the device does
          not have, the browser substitutes one, so line breaks can differ slightly from Office. Embedded fonts in the file
          are used where present.
        </p>
        <p>To make fallbacks deterministic, pin a chain:</p>
        <Code>{`<OfficeFile data={source} fallbackFonts={{ fallbackChain: ['Inter', 'Noto Sans'], cjkFallback: 'jp' }} />`}</Code>
        <p>
          {c('cjkFallback')} accepts {c("'sc' | 'tc' | 'hk' | 'jp' | 'kr'")} and puts that region's system CJK fonts first.
        </p>
      </Section>

      <Section id="low-level" title="Low-level API" lead="For hosts that parse or paint documents themselves. The components use these internally.">
        <H3>Parsing</H3>
        <Code lang="ts">{`import { OfficePackage, parseDocx, parseXlsx, parsePptx, parseOdt, parseRtf, isRtf } from 'officeview'

const pkg = await OfficePackage.load(bytes)   // the zip, with a cached part reader
const docx = await parseDocx(pkg)
const xlsx = await parseXlsx(pkg, { valuePolicy: 'engine' })   // or 'cached': saved values only
const pptx = await parsePptx(pkg)
const odt  = await parseOdt(pkg)              // a DocxDocument
const rtf  = parseRtf(bytes)                  // RTF is not a zip; isRtf(bytes) tests for it`}</Code>
        <H3>Painting</H3>
        <p>
          {c('getPaintables(doc, options?)')} turns any parsed document into page-sized units, each with a {c('spec')}{' '}
          ({c('widthPx')}, {c('heightPx')}) and a {c('paint(ctx)')} function. It prepares embedded fonts, images and charts first.
          Call {c('dispose()')} on the result when you no longer paint from it.
        </p>
        <Code lang="ts">{`import { getPaintables } from 'officeview'

const pages = await getPaintables(doc, { watermark: { text: 'DRAFT' } })
for (const page of pages) {
  const canvas = document.createElement('canvas')
  canvas.width = page.spec.widthPx * devicePixelRatio
  canvas.height = page.spec.heightPx * devicePixelRatio
  const ctx = canvas.getContext('2d')!
  ctx.scale(devicePixelRatio, devicePixelRatio)
  page.paint(ctx)
}
pages.dispose()`}</Code>
        <p>Format-specific functions are exported too:</p>
        <Table
          head={['Function', 'Purpose']}
          rows={[
            [c('layoutDocx(doc, measure)'), 'Line-break and paginate a Word, ODT or RTF document into PageLayout[].'],
            [c('renderPages(pages, ctx, images?)'), 'Paint laid-out pages.'],
            [c('computeMetrics(sheet) / renderSheet(sheet, ctx, metrics?)'), 'Measure and paint a worksheet.'],
            [c('slideMetrics(doc) / renderSlide(slide, ctx, metrics?, images?)'), <>Measure and paint a slide. Await {c('enrichSlideAppearance(doc)')} first for backgrounds and master artwork.</>],
            [c('buildTextIndex(paintables) / findMatches(index, query)'), 'The search index the viewer uses.'],
            [c('renderPageBitmap(worker, request)'), <>Render a page to an {c('ImageBitmap')} in a worker ({c('officeview/worker')}). {c('isWorkerRenderSupported()')} checks for OffscreenCanvas.</>],
            [c('protectionFromHeaders(headers) / watermarkFromHeaders(headers)'), 'Read the server headers yourself.'],
            [c('clampZoom, stepZoom, zoomAt, …'), 'The zoom maths behind the viewer, for custom controls.'],
          ]}
        />
      </Section>

      <Section id="types" title="Types">
        <p>All public types are exported from {c("'officeview'")}. The ones you are most likely to use:</p>
        <Code lang="ts">{`import type {
  OfficeFileProps, OfficeDocProps,           // component props
  OfficeViewerHandle, SlideShowSession,      // slideshow control
  SlideShowErrorCode, SlideShowDiagnostic,
  ByteSource, HttpSource, Progress,          // loading
  WatermarkOptions, ProtectionPolicy,        // protection
  DocxDocument, XlsxDocument, PptxDocument,  // parsed models
  Paintable, PaintableArray, PaintOptions,   // painting
} from 'officeview'`}</Code>
        <Code lang="ts">{`interface SlideShowDiagnostic {
  code: string
  slideIndex?: number
  effect?: string
  message: string
}`}</Code>
      </Section>

      <Section id="limitations" title="Known limitations">
        <ul>
          <li>Text is laid out with the browser's {c('measureText')}, without OpenType shaping, so line breaks can differ slightly from Word.</li>
          <li>Floating images are positioned from their anchor but text does not yet wrap around them.</li>
          <li>RTF skips vector metafiles ({c('\\emfblip')}, {c('\\wmetafile')}), drawing shapes, fields and embedded objects, and flattens nested tables.</li>
          <li>ODT does not render text boxes, vector shapes, multi-column text or nested tables.</li>
          <li>3-D transitions and Morph are close look-alikes of PowerPoint, not frame-exact copies. There is no pen or laser pointer.</li>
          <li>The interactive spreadsheet viewer is experimental; Safari, physical touch devices and VoiceOver have not been verified.</li>
          <li>Copy and print prevention and watermarks are deterrents, not DRM.</li>
        </ul>
        <p>
          The full list, with validation records, is in the <a href={`${REPO}#readme`}>README</a> and{' '}
          <a href={`${REPO}/tree/main/docs/validation`}>docs/validation</a>. Please report problems on{' '}
          <a href={`${REPO}/issues`}>GitHub</a>.
        </p>
      </Section>
    </>
  )
}

// ---------------------------------------------------------------------------

mount('docs', <DocShell nav={NAV} label="Documentation"><Content /></DocShell>)
