import { Code } from './code'
import { DocShell, H3, Note, Section, Table, c, type DocNav } from './doc-kit'
import { REPO, href, mount } from './layout'

const NAV: DocNav = [
  ['Framework setup', [['requirements', 'Requirements'], ['vite', 'Vite'], ['nextjs', 'Next.js'], ['webpack', 'webpack'], ['viewer-worker', 'The viewer worker']]],
  ['Serving documents', [['serving', 'HTTP checklist'], ['server-example', 'Server example'], ['csp', 'Content Security Policy']]],
  ['In production', [['performance', 'Performance'], ['accessibility', 'Accessibility'], ['testing', 'Testing your app'], ['troubleshooting', 'Troubleshooting']]],
  ['Contributing', [['architecture', 'Architecture'], ['repo-layout', 'Repository layout'], ['dev-setup', 'Development setup'], ['tests', 'Tests'], ['releasing', 'Releasing']]],
]

const tree = `src/
  components/   React components: OfficeFile, OfficeDoc, toolbar, interactive sheet viewer
  core/         zip reader, byte streams, fonts, search, selection, zoom, watermark
  docx/         Word parser and layout engine (pagination, line breaking, tables)
  odt/ odf/     OpenDocument Text parser, mapped onto the Word model
  rtf/          Rich Text parser, mapped onto the Word model
  xlsx/         spreadsheet parser, formula engine, grid renderer, interactive view
  pptx/         presentation parser, slide renderer, slide appearance
  drawing/      DrawingML shapes, pictures and preset geometry shared by all formats
  charts/       the chart engine (lazy)
  slideshow/    the show engine: timing, transitions, WebGL2 and Canvas 2D (lazy)
  render/       getPaintables: one paint closure per page, sheet or slide
  worker/       the viewer worker and the render-to-bitmap worker
  styles/       officeview.css (Tailwind, scoped to the theme root)
  demo/         the dev demo (bun run dev)
  debug/        harness pages for browser and performance tests
  testdata/     builders that write OOXML/ODF test documents
tests/          Vitest suites (jsdom + node-canvas) and Playwright specs
scripts/        corpus fetch/report, goldens, harnesses, site samples
site/           this website
examples/       standalone Vite and Next.js projects
docs/           architecture notes and validation records`

function Content() {
  return (
    <>
      <Section id="requirements" title="Requirements" lead="officeview runs in the browser. What your app needs:">
        <ul>
          <li><strong>React 18 or later</strong> (a peer dependency).</li>
          <li><strong>An ESM bundler</strong> that can import CSS from a package ({c("import 'officeview/style.css'")}) and code-split dynamic imports. Vite, webpack 5, Turbopack, Rollup and esbuild all can.</li>
          <li><strong>A sized container.</strong> The viewer scrolls its own pages; see <a href={href('docs/#sizing')}>Sizing the viewer</a>.</li>
          <li><strong>Browser APIs:</strong> Canvas 2D everywhere; WebGL2 for 3-D slideshow transitions (with a Canvas 2D fallback); module workers and OffscreenCanvas only for large spreadsheets.</li>
        </ul>
        <p>
          The heavy parts load on demand: the toolbar, the chart engine, slide backgrounds and artwork, the slideshow, the
          interactive spreadsheet grid and printing are separate chunks. A Word document with no charts never downloads
          the chart engine.
        </p>
      </Section>

      <Section id="vite" title="Vite">
        <p>Install the package, import the component and stylesheet, and set the worker format:</p>
        <Code title="vite.config.ts" lang="ts">{`import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // The viewer worker code-splits its lazy modules: it needs ES-module output.
  worker: { format: 'es' },
})`}</Code>
        <Code title="src/viewer-worker.ts" lang="ts">{`import 'officeview/worker'`}</Code>
        <Code title="src/App.tsx">{`import { OfficeFile } from 'officeview'
import 'officeview/style.css'

const createViewerWorker = () =>
  new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })

export function App({ file }: { file: File }) {
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <OfficeFile data={file} interactiveSheets viewerWorker={createViewerWorker} />
    </div>
  )
}`}</Code>
        <p>
          The worker setting only matters if you pass {c('viewerWorker')}. Without it Vite's default {c("'iife'")} worker format
          fails the build. The full project is in{' '}
          <a href={`${REPO}/tree/main/examples/vite-react`}>examples/vite-react</a>.
        </p>
      </Section>

      <Section id="nextjs" title="Next.js" lead="Tested with the App Router in Next.js 16, with Turbopack (the default) and with webpack. No next.config changes are needed.">
        <p>
          officeview uses React hooks and draws on {c('<canvas>')}, so render it from a client component. Server
          components can render that client component.
        </p>
        <Code title="app/DocumentViewer.tsx">{`'use client'

import { useMemo } from 'react'
import { OfficeFile } from 'officeview'
import 'officeview/style.css'

const createViewerWorker = () =>
  new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })

export function DocumentViewer({ url }: { url: string }) {
  const source = useMemo(() => ({ url }), [url])   // a new object would reload the file
  return <OfficeFile data={source} interactiveSheets viewerWorker={createViewerWorker} />
}`}</Code>
        <Code title="app/viewer-worker.ts" lang="ts">{`import 'officeview/worker'`}</Code>
        <Code title="app/page.tsx">{`import { DocumentViewer } from './DocumentViewer'

export default function Page() {
  return (
    <main style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <DocumentViewer url="/files/report.docx" />
    </main>
  )
}`}</Code>
        <p>
          Files in {c('public/')} are served from the site root, so {c("{ url: '/files/report.docx' }")} works for{' '}
          {c('public/files/report.docx')}. The full project is in <a href={`${REPO}/tree/main/examples/nextjs`}>examples/nextjs</a>.
        </p>
      </Section>

      <Section id="webpack" title="webpack">
        <p>
          webpack 5 needs no configuration. It bundles {c("new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })")}{' '}
          as a worker chunk, including the worker's lazy imports. Import the stylesheet with your CSS loader (for example{' '}
          {c('css-loader')} with {c('style-loader')} or {c('mini-css-extract-plugin')}).
        </p>
        <Note>
          <p>
            The worker entry should be a file in your app that imports {c("'officeview/worker'")}, as above. Pointing{' '}
            {c('new URL()')} straight at a package path is not resolved the same way by every bundler.
          </p>
        </Note>
      </Section>

      <Section id="viewer-worker" title="The viewer worker">
        <p>
          The worker is only used by the interactive spreadsheet viewer ({c('interactiveSheets')}) for workbooks that are too
          large to hold in memory. On desktop that means, roughly, more than 100,000 rows or cells, more than 32 MiB compressed, or more than
          6 MiB of sheet XML (half these on phones and tablets). Such a workbook is paged into browser storage and the grid reads
          bounded windows of it from the worker.
        </p>
        <ul>
          <li>Each large document gets its own worker, terminated when the document closes.</li>
          <li>Without a worker, large workbooks open in a compatibility mode, with a notice and no memory guarantee.</li>
          <li>Large workbooks show their saved formula values; they are not recalculated.</li>
          <li>Browser storage (OPFS, else IndexedDB) is used for the pages; a full or denied store is reported to the reader.</li>
        </ul>
      </Section>

      <Section id="serving" title="HTTP checklist" lead="When the viewer downloads a file itself ({ url } or a Request), these response details matter.">
        <Table
          head={['Header', 'Why']}
          rows={[
            [c('Content-Length'), 'Gives the progress bar a total. Without it, progress counts bytes only.'],
            [c('Content-Disposition'), <>Its {c('filename')} becomes the name in the toolbar (otherwise the last URL segment is used).</>],
            [c('Access-Control-Allow-Origin'), 'Required when the file is on another origin. Add Access-Control-Allow-Credentials for cookies.'],
            [c('Access-Control-Expose-Headers'), <>Cross-origin, list {c('Content-Disposition')}, {c('X-OfficeView-Protection')} and {c('X-OfficeView-Watermark')}, or the browser hides them from the viewer.</>],
            [c('X-OfficeView-Protection'), <>Optional: {c('no-copy')}, {c('no-print')}. See <a href={href('docs/#server-headers')}>Server headers</a>.</>],
            [c('X-OfficeView-Watermark'), 'Optional: a watermark the client cannot remove with props.'],
          ]}
        />
        <Note>
          <p>
            Office files are already zip-compressed. Serving them with {c('gzip')} or {c('br')} gains almost nothing and makes{' '}
            {c('Content-Length')} the compressed size, so the progress total no longer matches the bytes received.
          </p>
        </Note>
        <p>
          A non-2xx response fails the load with {c('HTTP <status> while downloading the document')}, which your{' '}
          {c('error')} renderer receives.
        </p>
      </Section>

      <Section id="server-example" title="Server example" lead="An Express route that serves a document to a signed-in user, with a per-user watermark and copy/print denied.">
        <Code lang="ts" title="server.ts">{`import express from 'express'
import { createReadStream, statSync } from 'node:fs'

const app = express()

app.get('/api/documents/:id', requireUser, (req, res) => {
  const path = documentPath(req.params.id)            // your lookup and access check
  const { size } = statSync(path)
  const mark = encodeURIComponent(\`\${req.user.email} · \${new Date().toISOString().slice(0, 10)}\`)

  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'Content-Length': String(size),
    'Content-Disposition': 'inline; filename="report.docx"',
    'X-OfficeView-Protection': 'no-copy, no-print',
    'X-OfficeView-Watermark': \`text=\${mark}; placement=tile; opacity=0.12\`,
    // Only needed when the app is on another origin:
    'Access-Control-Expose-Headers': 'Content-Disposition, X-OfficeView-Protection, X-OfficeView-Watermark',
  })
  createReadStream(path).pipe(res)
})`}</Code>
        <Code title="Client">{`const source = useMemo(() => ({ url: \`/api/documents/\${id}\`, credentials: 'include' as const }), [id])

<OfficeFile data={source} />`}</Code>
        <p>The server's protection and watermark win over the component's props, so a client cannot turn them off.</p>
      </Section>

      <Section id="csp" title="Content Security Policy" lead="What a strict CSP needs to allow, and what happens if it does not.">
        <Table
          head={['Directive', 'Needed for']}
          rows={[
            [c("img-src 'self' blob:"), 'Embedded pictures are decoded from blob URLs, and Print renders pages to blob images.'],
            [c("media-src 'self' blob:"), 'Audio and video on slides, played from blob URLs during a slideshow.'],
            [c("worker-src 'self'"), 'The viewer worker for large spreadsheets (served from your origin).'],
            [c("style-src 'self'"), <>Link {c('officeview/style.css')} as a stylesheet file. The Presenter View window inherits the page's CSP and loads the same stylesheet URL.</>],
          ]}
        />
        <p>
          Print isolation adds small {c('<style>')} elements at run time: the spreadsheet viewer's print rules, and a print guard
          in the page viewer when {c('style.css')} was not imported. A {c('style-src')} without {c("'unsafe-inline'")} blocks them.
          The page viewer's guard is unnecessary when you import {c('style.css')}; for the interactive spreadsheet viewer,
          check that {c('allowPrint={false}')} still hides the sheet from your print output under your policy.
        </p>
        <p>Inline {c('style')} props set through React are applied through the DOM and are not affected by {c('style-src')}.</p>
      </Section>

      <Section id="performance" title="Performance">
        <ul>
          <li><strong>Let the lazy chunks stay lazy.</strong> Import from {c("'officeview'")} only; the charts, slideshow, grid and toolbar chunks load when a document needs them.</li>
          <li><strong>Warm the chart engine</strong> when you know charts are coming (for example while the file downloads): {c('preloadCharts()')} returns a promise that resolves once the chunk is loaded.</li>
          <li><strong>Reuse the parsed model.</strong> If you show the same file twice, parse it once with {c('loadOfficeFile')} and pass it to several {c('<OfficeDoc>')}s.</li>
          <li><strong>Dispose headless paint results.</strong> {c('getPaintables()')} holds fonts and decoded images until you call {c('dispose()')}.</li>
          <li><strong>Large spreadsheets</strong> belong in the worker: pass {c('viewerWorker')} with {c('interactiveSheets')}.</li>
          <li>Canvas backing stores follow {c('devicePixelRatio')} (capped at 3), and only visible pages re-render when zoomed.</li>
        </ul>
      </Section>

      <Section id="accessibility" title="Accessibility">
        <ul>
          <li>The page region is a focusable {c('role="region"')} labelled “Document viewer”. Find, zoom, Present and the toolbar are keyboard operable, and touch targets are 44 px on coarse pointers.</li>
          <li>Charts are exposed to assistive technology as images labelled with a description of their data.</li>
          <li>The interactive spreadsheet grid is a single Tab stop with arrow-key navigation, announced selection and keyboard sorting, filtering and resizing.</li>
          <li>The slideshow follows {c('prefers-reduced-motion')} unless you set {c('slideShowMotion')}.</li>
          <li>
            Text is painted on canvas, so screen readers cannot read the page body the way they read HTML. If your readers
            need the text, offer the original file for download as well.
          </li>
        </ul>
      </Section>

      <Section id="testing" title="Testing your app">
        <p>
          The viewer needs a real Canvas 2D context. jsdom has none unless the {c('canvas')} npm package is installed (that is how
          officeview's own suite runs). For your app's unit tests, mock the component and test your own logic:
        </p>
        <Code lang="ts" title="viewer.test.tsx">{`import { vi } from 'vitest'

vi.mock('officeview', () => ({
  OfficeFile: (props: { data: unknown }) => <div data-testid="office-file" />,
}))`}</Code>
        <p>
          Test the real rendering in a browser with Playwright or Cypress: wait for the region labelled “Document viewer” and
          assert on the status line, or compare screenshots.
        </p>
      </Section>

      <Section id="troubleshooting" title="Troubleshooting">
        <Table
          head={['Symptom', 'Cause and fix']}
          rows={[
            ['The document reloads or flickers when the parent re-renders', <>{c('data')} is compared by identity. Create the source once with {c('useMemo')} instead of an inline {c('{{ url }}')}.</>],
            [<>Unstyled toolbar; console says {c('style.css not loaded')}</>, <>Import {c("'officeview/style.css'")} once in your app.</>],
            ['The page counter stays at 1 / N', 'The viewer is not in a sized container, so the page scrolls instead of the viewer. See Sizing the viewer.'],
            [<>Vite build error about the worker format</>, <>Set {c("worker: { format: 'es' }")} in {c('vite.config')}.</>],
            [<>{c('present()')} rejects with {c('gesture-required')}</>, <>Call it synchronously in a click or key handler, with no {c('await')} before it.</>],
            ['Server watermark or protection ignored cross-origin', <>Add the headers to {c('Access-Control-Expose-Headers')}.</>],
            ['Text wraps differently from Word', 'A font named by the document is missing on the device. Install it, or pin a fallback chain with fallbackFonts.'],
            [<>{c('HTTP 401 while downloading the document')}</>, <>Pass {c('headers')} or {c("credentials: 'include'")} in the {c('{ url }')} descriptor.</>],
          ]}
        />
      </Section>

      <Section id="architecture" title="Architecture" lead="Each format is parsed straight from its XML into a typed model, laid out, then painted to canvas.">
        <Code lang="bash" title="Pipeline">{`bytes ─▶ OfficePackage (zip, cached part reader)
          ├─▶ parseDocx / parseOdt / parseRtf ─▶ layoutDocx (line breaking, pagination) ─▶ renderPages
          ├─▶ parseXlsx ─▶ computeMetrics (grid) ─▶ renderSheet   |  interactive grid (tiles, worker)
          └─▶ parsePptx ─▶ slideMetrics (EMU) ─▶ renderSlide      |  slideshow engine (WebGL2 / 2D)
                              │
              getPaintables(doc) ─▶ one paint(ctx) per page ─▶ <OfficeDoc> canvases, print, search, thumbnails`}</Code>
        <ul>
          <li><strong>One model per family.</strong> ODT and RTF map onto the Word model, so pagination, search, selection and zoom come from one engine.</li>
          <li><strong>Paint closures.</strong> {c('getPaintables')} returns a {c('paint(ctx)')} per page. The viewer, print, the search index (by replaying paint into a recording context) and the worker all use them.</li>
          <li><strong>Shared DrawingML.</strong> Shapes, pictures, preset geometry and charts are parsed and painted by the same code in all three OOXML formats.</li>
          <li><strong>Budgets everywhere.</strong> Parsers, images, charts and the grid run under explicit memory and size limits; anything refused is diagnosed, never silently dropped.</li>
        </ul>
        <p>Design notes for the spreadsheet viewer are in <a href={`${REPO}/tree/main/docs/architecture`}>docs/architecture</a>.</p>
      </Section>

      <Section id="repo-layout" title="Repository layout">
        <Code lang="bash" title="officeview/">{tree}</Code>
      </Section>

      <Section id="dev-setup" title="Development setup">
        <p>
          The repo uses <a href="https://bun.sh">bun</a> (CI pins 1.3.10). Tests render with node-canvas, which needs Cairo and
          Pango; on Debian or Ubuntu:
        </p>
        <Code lang="bash">{`sudo apt-get install -y libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev   # macOS: brew install pkg-config cairo pango libpng jpeg giflib librsvg
git clone https://github.com/julianshen/officeview && cd officeview
bun install
bun run dev          # demo app at http://localhost:5173: open files or generated samples
bun run site:dev     # this website at http://localhost:5174/officeview/`}</Code>
        <Table
          head={['Command', 'What it does']}
          rows={[
            [c('bun run test'), 'The Vitest suite (jsdom + node-canvas, pixel-sampled assertions).'],
            [c('bun run build'), 'Type-check, library build to dist/, and declarations.'],
            [c('bun run lint'), <>{c('tsc --noEmit')}.</>],
            [c('bun run test:browser'), 'Playwright specs in Chromium (and WebKit), including the slideshow smoke tests.'],
            [c('bun scripts/fetch-corpus.ts'), 'Download the real-file corpus, pinned to upstream commits and verified by SHA-256.'],
            [c('bun scripts/corpus-report.ts'), 'Run every corpus file through the pipeline and fail on degraded output.'],
            [c('bun run site:build'), 'Build this website into site-dist/.'],
          ]}
        />
      </Section>

      <Section id="tests" title="Tests">
        <H3>What CI runs</H3>
        <p>
          On every push and pull request, <a href={`${REPO}/blob/main/.github/workflows/ci.yml`}>ci.yml</a> type-checks, fetches
          the corpus, runs the suite without the pixel-golden files, builds, and runs the corpus report. A second job runs the
          slideshow Playwright project on SwiftShader WebGL2. <a href={`${REPO}/blob/main/.github/workflows/pages.yml`}>pages.yml</a>{' '}
          builds this site.
        </p>
        <Code lang="bash">{`bun scripts/fetch-corpus.ts
bunx vitest run --exclude 'tests/golden*.test.ts'`}</Code>
        <H3>Golden images</H3>
        <p>
          Golden PNGs are pixel-exact and depend on the machine's font rasteriser, so CI skips them. Compare them locally on the
          machine that recorded them; mismatches write {c('actual')} and red-mask {c('diff')} images to{' '}
          {c('/tmp/officeview-diff/')}.
        </p>
        <Code lang="bash">{`bun scripts/golden-corpus.ts compare      # real corpus files
bun scripts/golden.ts compare <file.docx> --name my-doc`}</Code>
        <H3>Writing tests</H3>
        <p>
          Build test documents with the builders in {c('src/testdata')} ({c('buildDocx')}, {c('buildXlsx')},{' '}
          {c('buildPptxDeck')}, {c('buildOdt')}, the chart gallery) rather than committing binaries. Real-world regressions go in the
          corpus, which is fetched rather than vendored.
        </p>
      </Section>

      <Section id="releasing" title="Releasing">
        <ol>
          <li>Bump {c('version')} in {c('package.json')} and move the {c('Unreleased')} changelog entries under the new version.</li>
          <li>Open a pull request and merge it once CI is green.</li>
          <li>
            On GitHub, draft a release with a new tag {c('vX.Y.Z')} on {c('main')} and publish it. The{' '}
            <a href={`${REPO}/blob/main/.github/workflows/publish.yml`}>publish workflow</a> checks that the tag matches{' '}
            {c('package.json')}, builds, and publishes to npm with provenance.
          </li>
        </ol>
        <Note>
          <p>
            Publishing uses the {c('NPM_TOKEN')} repository secret. Re-running the workflow for a version that is already on npm
            does nothing, so it is safe to retry.
          </p>
        </Note>
        <p>Every push to {c('main')} also redeploys this website.</p>
      </Section>
    </>
  )
}

mount('guides', <DocShell nav={NAV} label="Guides"><Content /></DocShell>)
