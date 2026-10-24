import { useEffect, useMemo, useRef, useState, type DragEvent, type ReactNode } from 'react'
import { SlideShowError, type OfficeViewerHandle, type WatermarkOptions } from 'officeview'
import { Code } from './code'
import { Icon, mount, useSiteTheme } from './layout'
import { SAMPLES, sampleUrl } from './samples'
import { Viewer } from './viewer'
import './playground.css'

type Theme = 'site' | 'system' | 'light' | 'dark'
type Placement = NonNullable<WatermarkOptions['placement']>
type Motion = 'auto' | 'full' | 'reduced'

interface Options {
  theme: Theme
  showToolbar: boolean
  showZoomControls: boolean
  showSearch: boolean
  allowCopy: boolean
  allowPrint: boolean
  allowPresent: boolean
  presentShortcuts: boolean
  slideShowMotion: Motion
  interactiveSheets: boolean
  pageGapPx: number
  watermarkOn: boolean
  watermark: { text: string; placement: Placement; rotate: number; opacity: number; color: string }
}

const DEFAULTS: Options = {
  theme: 'site',
  showToolbar: true,
  showZoomControls: true,
  showSearch: true,
  allowCopy: true,
  allowPrint: true,
  allowPresent: true,
  presentShortcuts: true,
  slideShowMotion: 'auto',
  interactiveSheets: false,
  pageGapPx: 16,
  watermarkOn: false,
  watermark: { text: 'CONFIDENTIAL', placement: 'tile', rotate: -45, opacity: 0.15, color: '#000000' },
}

type Source = { kind: 'sample'; id: string } | { kind: 'file'; file: File }

const ACCEPT = '.docx,.xlsx,.pptx,.odt,.rtf'

function initialSource(): Source {
  const id = location.hash.slice(1)
  return { kind: 'sample', id: SAMPLES.some((s) => s.id === id) ? id : 'report' }
}

/** The JSX a host would write for the current options (defaults omitted). */
function snippet(o: Options, name: string): string {
  const attrs: string[] = []
  const add = (s: string) => attrs.push(s)
  if (o.theme !== 'site' && o.theme !== 'system') add(`theme="${o.theme}"`)
  if (!o.showToolbar) add('showToolbar={false}')
  if (!o.showZoomControls) add('showZoomControls={false}')
  if (!o.showSearch) add('showSearch={false}')
  if (!o.allowCopy) add('allowCopy={false}')
  if (!o.allowPrint) add('allowPrint={false}')
  if (!o.allowPresent) add('allowPresent={false}')
  if (!o.presentShortcuts) add('presentShortcuts={false}')
  if (o.slideShowMotion !== 'auto') add(`slideShowMotion="${o.slideShowMotion}"`)
  if (o.interactiveSheets) add('interactiveSheets')
  if (o.pageGapPx !== 16) add(`pageGapPx={${o.pageGapPx}}`)
  if (o.watermarkOn && o.watermark.text) {
    const w = o.watermark
    const parts = [`text: '${w.text.replace(/'/g, "\\'")}'`]
    if (w.placement !== 'tile') parts.push(`placement: '${w.placement}'`)
    if (w.rotate !== -45) parts.push(`rotate: ${w.rotate}`)
    if (w.opacity !== 0.15) parts.push(`opacity: ${w.opacity}`)
    if (w.color !== '#000000') parts.push(`color: '${w.color}'`)
    add(`watermark={{ ${parts.join(', ')} }}`)
  }
  const data = 'data={source}'
  const all = [data, ...attrs]
  const body = all.length <= 2 && all.join(' ').length < 60 ? ` ${all.join(' ')} />` : `\n${all.map((a) => `  ${a}`).join('\n')}\n/>`
  return `import { OfficeFile } from 'officeview'
import 'officeview/style.css'

const source = { url: '/files/${name}' }   // create once: a new object reloads the file

<OfficeFile${body}`
}

function Field({ label, hint, children }: { label: string; hint?: string; children: ReactNode }) {
  return (
    <label className="field">
      <span className="field-label">{label}{hint && <span className="hint">{hint}</span>}</span>
      {children}
    </label>
  )
}

function Toggle({ label, hint, checked, onChange }: { label: string; hint?: string; checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <label className="toggle">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span className="switch" aria-hidden="true" />
      <span className="toggle-label">{label}{hint && <span className="hint">{hint}</span>}</span>
    </label>
  )
}

function Playground() {
  const [source, setSource] = useState<Source>(initialSource)
  const [o, setO] = useState<Options>(DEFAULTS)
  const [dragging, setDragging] = useState(false)
  const [log, setLog] = useState<string[]>([])
  const [siteTheme] = useSiteTheme()
  const viewer = useRef<OfficeViewerHandle>(null)
  const fileInput = useRef<HTMLInputElement>(null)
  const set = <K extends keyof Options>(k: K, v: Options[K]) => setO((prev) => ({ ...prev, [k]: v }))
  const setMark = <K extends keyof Options['watermark']>(k: K, v: Options['watermark'][K]) => setO((prev) => ({ ...prev, watermark: { ...prev.watermark, [k]: v } }))

  const sample = source.kind === 'sample' ? SAMPLES.find((s) => s.id === source.id)! : null
  const name = sample ? sample.file : source.kind === 'file' ? source.file.name : ''
  const src = useMemo(() => (source.kind === 'sample' ? sampleUrl(SAMPLES.find((s) => s.id === source.id)!) : source.file), [source])
  const isDeck = /\.pptx$/i.test(name)
  const isSheet = /\.xlsx$/i.test(name)

  useEffect(() => {
    // Samples that need an option to show their point turn it on.
    if (sample?.props?.interactiveSheets) set('interactiveSheets', true)
  }, [sample?.id]) // eslint-disable-line react-hooks/exhaustive-deps

  const note = (line: string) => setLog((l) => [`${new Date().toLocaleTimeString()}  ${line}`, ...l].slice(0, 40))

  const openFile = (file: File | undefined) => {
    if (!file) return
    setSource({ kind: 'file', file })
    history.replaceState(null, '', location.pathname)
    note(`Opened ${file.name} (${(file.size / 1024).toFixed(1)} KiB) from this device`)
  }
  const onDrop = (e: DragEvent) => {
    e.preventDefault()
    setDragging(false)
    openFile(e.dataTransfer.files[0])
  }

  const theme = o.theme === 'site' ? siteTheme : o.theme
  const watermark = o.watermarkOn && o.watermark.text ? o.watermark : undefined

  return (
    <main className="playground">
      <aside className="panel" aria-label="Viewer options">
        <section>
          <h2>Document</h2>
          <Field label="Sample">
            <select value={source.kind === 'sample' ? source.id : ''} onChange={(e) => { setSource({ kind: 'sample', id: e.target.value }); history.replaceState(null, '', `#${e.target.value}`) }}>
              {source.kind === 'file' && <option value="">{source.file.name}</option>}
              {SAMPLES.map((s) => <option key={s.id} value={s.id}>{s.title} ({s.format})</option>)}
            </select>
          </Field>
          <button className="btn btn-secondary btn-sm upload" onClick={() => fileInput.current?.click()}><Icon.upload /> Open a file…</button>
          <input ref={fileInput} type="file" accept={ACCEPT} hidden onChange={(e) => { openFile(e.target.files?.[0]); e.target.value = '' }} />
          <p className="privacy"><Icon.shield /> Files are read in your browser and never uploaded.</p>
        </section>

        <section>
          <h2>Appearance</h2>
          <Field label="theme">
            <select value={o.theme} onChange={(e) => set('theme', e.target.value as Theme)}>
              <option value="site">Follow this site</option>
              <option value="system">system</option>
              <option value="light">light</option>
              <option value="dark">dark</option>
            </select>
          </Field>
          <Toggle label="showToolbar" checked={o.showToolbar} onChange={(v) => set('showToolbar', v)} />
          <Toggle label="showZoomControls" checked={o.showZoomControls} onChange={(v) => set('showZoomControls', v)} />
          <Toggle label="showSearch" checked={o.showSearch} onChange={(v) => set('showSearch', v)} />
          <Field label="pageGapPx" hint={`${o.pageGapPx}px`}>
            <input type="range" min={0} max={48} step={2} value={o.pageGapPx} onChange={(e) => set('pageGapPx', +e.target.value)} />
          </Field>
        </section>

        <section>
          <h2>Protection</h2>
          <Toggle label="allowCopy" checked={o.allowCopy} onChange={(v) => set('allowCopy', v)} />
          <Toggle label="allowPrint" checked={o.allowPrint} onChange={(v) => set('allowPrint', v)} />
          <Toggle label="watermark" checked={o.watermarkOn} onChange={(v) => set('watermarkOn', v)} />
          {o.watermarkOn && (
            <div className="sub">
              <Field label="text"><input type="text" value={o.watermark.text} onChange={(e) => setMark('text', e.target.value)} /></Field>
              <Field label="placement">
                <select value={o.watermark.placement} onChange={(e) => setMark('placement', e.target.value as Placement)}>
                  {(['tile', 'center', 'header', 'footer'] as const).map((p) => <option key={p}>{p}</option>)}
                </select>
              </Field>
              <Field label="rotate" hint={`${o.watermark.rotate}°`}>
                <input type="range" min={-90} max={90} step={5} value={o.watermark.rotate} onChange={(e) => setMark('rotate', +e.target.value)} />
              </Field>
              <Field label="opacity" hint={o.watermark.opacity.toFixed(2)}>
                <input type="range" min={0.05} max={0.6} step={0.01} value={o.watermark.opacity} onChange={(e) => setMark('opacity', +e.target.value)} />
              </Field>
              <Field label="color"><input type="color" value={o.watermark.color} onChange={(e) => setMark('color', e.target.value)} /></Field>
            </div>
          )}
        </section>

        <section className={isSheet ? '' : 'dim'}>
          <h2>Spreadsheets</h2>
          <Toggle label="interactiveSheets" hint="experimental" checked={o.interactiveSheets} onChange={(v) => set('interactiveSheets', v)} />
        </section>

        <section className={isDeck ? '' : 'dim'}>
          <h2>Slideshow</h2>
          <Toggle label="allowPresent" checked={o.allowPresent} onChange={(v) => set('allowPresent', v)} />
          <Toggle label="presentShortcuts" hint="F5" checked={o.presentShortcuts} onChange={(v) => set('presentShortcuts', v)} />
          <Field label="slideShowMotion">
            <select value={o.slideShowMotion} onChange={(e) => set('slideShowMotion', e.target.value as Motion)}>
              <option>auto</option><option>full</option><option>reduced</option>
            </select>
          </Field>
          <button
            className="btn btn-primary btn-sm"
            disabled={!isDeck || !o.allowPresent}
            onClick={() => {
              viewer.current?.present({ from: 'beginning' })
                .then((session) => { note(`present() resolved: surface=${session.surface}, presenter=${session.presenter}`); return session.ended })
                .then(() => note('session.ended resolved'))
                .catch((e) => note(e instanceof SlideShowError ? `present() rejected: ${e.code}` : `present() failed: ${String(e)}`))
            }}
          >
            <Icon.play /> present()
          </button>
        </section>

        <button className="reset" onClick={() => setO(DEFAULTS)}>Reset all options</button>
      </aside>

      <div className="stage">
        <div
          className={`drop${dragging ? ' over' : ''}`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true) }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
        >
          <Viewer
            key={source.kind === 'file' ? `file:${source.file.name}:${source.file.lastModified}` : source.id}
            ref={viewer}
            src={src}
            height="calc(100vh - var(--header-h) - 48px)"
            frameStyle={{ minHeight: 480 }}
            theme={theme}
            showToolbar={o.showToolbar}
            showZoomControls={o.showZoomControls}
            showSearch={o.showSearch}
            allowCopy={o.allowCopy}
            allowPrint={o.allowPrint}
            allowPresent={o.allowPresent}
            presentShortcuts={o.presentShortcuts}
            slideShowMotion={o.slideShowMotion}
            interactiveSheets={o.interactiveSheets}
            pageGapPx={o.pageGapPx}
            watermark={watermark}
            onSlideShowStart={() => note('onSlideShowStart')}
            onSlideShowEnd={() => note('onSlideShowEnd')}
            onSlideShowSlideChange={(slide, build) => note(`onSlideShowSlideChange(slide ${slide}, build ${build})`)}
            onSlideShowDiagnostic={(d) => note(`onSlideShowDiagnostic: ${JSON.stringify(d)}`)}
          />
          {dragging && <div className="drop-hint"><Icon.upload /> Drop to open</div>}
        </div>
        <div className="outputs">
          <div>
            <h2>Code</h2>
            <Code title="Viewer.tsx">{snippet(o, name)}</Code>
          </div>
          <div>
            <h2>Events</h2>
            <div className="log" role="log" aria-live="polite">
              {log.length ? log.map((l, i) => <div key={i}>{l}</div>) : <span className="empty">Slideshow callbacks and present() results appear here.</span>}
            </div>
          </div>
        </div>
      </div>
    </main>
  )
}

mount('playground', <Playground />, { footer: false })
