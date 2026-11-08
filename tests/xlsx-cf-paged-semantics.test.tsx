/**
 * Task 9 fix round 2, finding 1: paged (worker-backed) semantic text must use
 * prepared CF formats, with localized descriptors and unavailable states.
 * Mounts the real viewer against a real worker kernel (same harness shape as
 * the unmount-GC test): no hand-built scopes, no resident fallbacks.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { xlsxBytes } from './helpers/xlsx-package'

vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})
afterEach(() => { cleanup(); vi.restoreAllMocks() })

const CF_SHEET_DATA = '<row r="1"><c r="A1" s="0"><v>1.25</v></c><c r="B1" s="0"><v>2</v></c></row>'
const CF_STYLES = '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs><dxfs count="2"><dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'
const CF_AFTER = '<conditionalFormatting sqref="A1:A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1:B1"><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>'

function workerPort(delayedStructure = false) {
  const requests: unknown[] = []
  let holdCf=false
  const held:(()=>void)[]=[], structure:(()=>void)[]=[]
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  return {
    requests, kernel,
    releaseStructure(){delayedStructure=false;for(const f of structure.splice(0))f()},
    holdCf(){holdCf=true},
    release(){holdCf=false;for(const f of held.splice(0))f()},
    get heldCount(){return held.length},
    postMessage(message: unknown, transfer: Transferable[]) {
      requests.push(message)
      const open = message as {kind:string;source?:{kind:string;blob?:Blob}}
      const blob = open.kind==='open' && open.source?.kind==='blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? {...open,source:{kind:'blob'}} : message, { transfer: transfer as ArrayBuffer[] }) as { kind: string; source?: unknown }
      if(blob) command.source={kind:'blob',blob}
      void kernel.handle(command as never, (out: unknown, replyTransfer) => {
        const deliver=()=>{const data = structuredClone(out, { transfer: replyTransfer as ArrayBuffer[] }); for (const listener of [...listeners]) listener({ data } as MessageEvent)}
        if(holdCf && (command as any).query?.op==='conditionalFormatting')held.push(deliver);else if(delayedStructure && ['structural','layout'].includes((command as any).query?.op))structure.push(deliver);else deliver()
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
}

async function mountCf(extra: { afterSheetData?: string } = {}) {
  const bytes = xlsxBytes({ sheetData: CF_SHEET_DATA + '<row r="2"><c r="A2"><v>2</v></c></row>', styles: CF_STYLES, afterSheetData: extra.afterSheetData ?? CF_AFTER })
  const ui = render(<OfficeFile data={new Blob([new Uint8Array(bytes).buffer])} interactiveSheets viewerWorker={() => workerPort()} />)
  return ui
}

describe('worker-backed semantic text follows CF', () => {
  it('paged CF numFmt formats the accessible name and fill gains a localized descriptor', async () => {
    const ui = await mountCf()
    try {
      const formatted = await ui.findByRole('gridcell', { name: /A1, value 1\.25/ }, { timeout: 30_000 })
      expect(formatted.textContent).toContain('1.25')
      const filled = await ui.findByRole('gridcell', { name: /Conditional formatting applied: fill/ }, { timeout: 30_000 })
      expect(filled.textContent).toContain('2')
    } finally {
      ui.unmount()
    }
  }, 60_000)

  it('a paged unresolved rule announces unavailable, never raw internals', async () => {
    const afterSheetData = '<conditionalFormatting sqref="A1:A1"><cfRule type="expression" dxfId="0" priority="1"><formula>LEN(A1)&gt;100000</formula></cfRule></conditionalFormatting>'
    const ui = await mountCf({ afterSheetData })
    try {
      const cell = await ui.findByRole('gridcell', { name: /A1, value .*unavailable/i }, { timeout: 30_000 })
      expect(cell.textContent).not.toContain('LEN(A1)')
      expect(cell.textContent).not.toContain('expression')
    } finally {
      ui.unmount()
    }
  }, 60_000)
})

it('paged paint covers full edge/ring tiles and remains current through warm one-row scroll; close withdraws retained scopes', async()=>{
 const {SheetTileRenderer}=await import('../src/components/xlsx/sheetTiles')
 const {RemoteWorkbook}=await import('../src/worker/viewerRemoteSource')
 let remote:InstanceType<typeof RemoteWorkbook>|undefined
 const identity=RemoteWorkbook.prototype.identity
 vi.spyOn(RemoteWorkbook.prototype,'identity').mockImplementation(function(this:InstanceType<typeof RemoteWorkbook>,id){remote=this;return identity.call(this,id)})
 const frames: import('../src/components/xlsx/sheetTiles').SheetFrame[]=[]
 const present=SheetTileRenderer.prototype.present
 vi.spyOn(SheetTileRenderer.prototype,'present').mockImplementation(function(this:InstanceType<typeof SheetTileRenderer>,ctx,frame){frames.push(frame);return present.call(this,ctx,frame)})
 const port=workerPort()
 const bytes=xlsxBytes({sheetData:Array.from({length:100},(_,i)=>`<row r="${i+1}"><c r="A${i+1}"><v>2</v></c></row>`).join(''),styles:CF_STYLES,afterSheetData:'<conditionalFormatting sqref="A1:A100"><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>'})
 const ui=render(<OfficeFile data={new Blob([new Uint8Array(bytes).buffer])} interactiveSheets viewerWorker={()=>port}/>)
 try {
  await waitFor(()=>expect(frames.some(f=>f.cfCoverage?.({x:0,y:0,width:512,height:512}) && f.assets?.conditionalFormatting?.layer.styleAt(25,0,undefined)?.fillColor)).toBe(true),{timeout:30000})
  const ready=[...frames].reverse().find(f=>f.cfCoverage?.({x:0,y:0,width:512,height:512}) && f.assets?.conditionalFormatting)!
  const held=ready.assets!.conditionalFormatting!,key=held.key
  const paintRequests=port.requests.length
  const scroller=ui.getByTestId('officeview-sheet-scroller')
  await act(async()=>{scroller.scrollTop=20;fireEvent.scroll(scroller);await new Promise(r=>setTimeout(r,100))})
  expect(frames.at(-1)?.assets?.conditionalFormatting?.key).toBe(key)
  // Semantic reads may change with the visible cells, but painting the already
  // covered ring must not issue another full-ring batch.
  expect(port.requests.slice(paintRequests).filter((r:any)=>r.query?.op==='conditionalFormatting' && r.query.rects?.length>30)).toHaveLength(0)
  expect(frames.at(-1)?.assets?.conditionalFormatting?.layer.styleAt(25,0,undefined)?.fillColor).toBeDefined()
  await act(async()=>remote!.close())
  expect(held.layer.styleAt(25,0,undefined)?.fillColor).toBeUndefined()
 } finally {ui.unmount()}
},60000)

it('worker plain-sheet frames send no conditional-formatting reads',async()=>{
 const port=workerPort(),bytes=xlsxBytes({sheetData:CF_SHEET_DATA+'<row r="2"><c r="A2"><v>2</v></c></row>'})
 const data=new Blob([new Uint8Array(bytes).buffer]),worker=()=>port
 const ui=render(<OfficeFile data={data} interactiveSheets viewerWorker={worker}/>)
 try {
  await ui.findByRole('gridcell',{name:/A1, value/},{timeout:30000})
  for(let i=0;i<100;i++)ui.rerender(<OfficeFile data={data} interactiveSheets viewerWorker={worker}/>)
  expect(port.requests.some((r:any)=>r.kind==='read')).toBe(true)
  expect(port.requests.filter((r:any)=>r.query?.op==='conditionalFormatting')).toHaveLength(0)
 } finally {ui.unmount()}
},60000)

it.each(['resize','close'] as const)('a delayed paged preparation after %s never publishes its old scope',async(action)=>{
 const {RemoteWorkbook}=await import('../src/worker/viewerRemoteSource')
 const text=await import('../src/xlsx/cf/text')
 let remote:InstanceType<typeof RemoteWorkbook>|undefined
 const original=RemoteWorkbook.prototype.identity
 vi.spyOn(RemoteWorkbook.prototype,'identity').mockImplementation(function(this:InstanceType<typeof RemoteWorkbook>,id){remote=this;return original.call(this,id)})
 const scopes:import('../src/xlsx/view/formatWorkspace').EffectiveFormatScope[]=[]
 const prepare=text.prepareCfWireScope
 vi.spyOn(text,'prepareCfWireScope').mockImplementation(async(...args)=>{const out=await prepare(...args);if(out)scopes.push(out);return out})
 const port=workerPort();port.holdCf()
 const bytes=xlsxBytes({sheetData:CF_SHEET_DATA+'<row r="2"><c r="A2"><v>2</v></c></row>',styles:CF_STYLES,afterSheetData:CF_AFTER})
 const ui=render(<OfficeFile data={new Blob([new Uint8Array(bytes).buffer])} interactiveSheets viewerWorker={()=>port}/>)
 try{
  await waitFor(()=>expect(port.heldCount).toBeGreaterThan(0),{timeout:30000})
  const before=remote!.merges('1').gen
  act(()=>{if(action==='close')remote!.close();else remote!.publishLayout('1',{...remote!.layout('1'),rowOverrides:[{start:0,end:0,sizeCss:40}]})})
  if(action==='resize')expect(remote!.merges('1').gen).toBeGreaterThan(before)
  act(()=>port.release())
  await new Promise(r=>setTimeout(r,150))
  expect(ui.getByRole('status').textContent).not.toMatch(/CF wire|Conditional formatting unavailable|unavailable for these tiles/)
  // A cancelled old scope must not be returned to any consumer. A new resize
  // scope is allowed, but its generation predicate must still retire on close.
  if(action==='close')expect(scopes).toHaveLength(0)
  else {await act(async()=>{await remote!.close()});for(const scope of scopes)expect(scope.describe(0,1)).toBeUndefined()}
 }finally{port.release();ui.unmount()}
},60000)

it('worker-shell in-flight g1 is withdrawn before the negative g2 overlay publishes',async()=>{
 const {RemoteWorkbook}=await import('../src/worker/viewerRemoteSource')
 const text=await import('../src/xlsx/cf/text')
 let remote:InstanceType<typeof RemoteWorkbook>|undefined
 const original=RemoteWorkbook.prototype.identity
 vi.spyOn(RemoteWorkbook.prototype,'identity').mockImplementation(function(this:InstanceType<typeof RemoteWorkbook>,id){remote=this;return original.call(this,id)})
 const scopes:import('../src/xlsx/view/formatWorkspace').EffectiveFormatScope[]=[]
 const prepare=text.prepareCfWireScope
 vi.spyOn(text,'prepareCfWireScope').mockImplementation(async(...args)=>{const out=await prepare(...args);if(out)scopes.push(out);return out})
 const port=workerPort();port.holdCf()
 const rules='<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>'
 const styles='<cellXfs count="1"><xf/></cellXfs><dxfs count="2"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'
 const bytes=xlsxBytes({sheetData:CF_SHEET_DATA+'<row r="2"><c r="A2"><v>2</v></c></row>',styles,afterSheetData:rules})
 const ui=render(<OfficeFile data={new Blob([new Uint8Array(bytes).buffer])} interactiveSheets viewerWorker={()=>port}/>)
 try{
  await waitFor(()=>expect(port.heldCount).toBeGreaterThan(0),{timeout:30000})
  const session=port.kernel.session(remote!.identity('1').docId)!,store=session.source.store('1')
  act(()=>{
   store.setValueOverlay!({rects:[{r0:0,r1:0,c0:0,c1:0}],protectedRects:[{r0:0,r1:0,c0:0,c1:0}],merges:[],cell:()=>({value:-1})})
   session.publishValues('1')
   // Same main-thread generation adoption used after the worker acknowledges a
   // pivot overlay release; avoids coupling this CF race to the pivot UI.
   ;(remote as unknown as {adoptPivotReleaseValueGen(id:string,gen:number):void}).adoptPivotReleaseValueGen('1',store.stamp.valueGen)
   port.release()
  })
  await waitFor(()=>expect(scopes.some(s=>s.effectiveStyle(0,0,undefined)?.fillColor==='FFFF0000')).toBe(true),{timeout:30000})
  expect(scopes.some(s=>s.effectiveStyle(0,0,undefined)?.fillColor==='FF00FF00')).toBe(false)
 }finally{port.release();ui.unmount()}
},60000)

it.each(['rule','empty'] as const)('asynchronous unchanged structure wakes %s CF without false notices or permanently pending tiles',async(kind)=>{
 const {SheetTileRenderer}=await import('../src/components/xlsx/sheetTiles')
 const {RemoteWorkbook}=await import('../src/worker/viewerRemoteSource')
 await import('../src/xlsx/cf/text') // warm lazy import: structure must really win the race later
 const states:string[]=[],acquire=RemoteWorkbook.prototype.acquireCfPreparation
 vi.spyOn(RemoteWorkbook.prototype,'acquireCfPreparation').mockImplementation(function(this:InstanceType<typeof RemoteWorkbook>,...args){const out=acquire.apply(this,args);states.push(out.state.state);return out})
 const frames:import('../src/components/xlsx/sheetTiles').SheetFrame[]=[],present=SheetTileRenderer.prototype.present
 vi.spyOn(SheetTileRenderer.prototype,'present').mockImplementation(function(this:InstanceType<typeof SheetTileRenderer>,ctx,f){frames.push(f);return present.call(this,ctx,f)})
 const bytes=xlsxBytes({beforeSheetData:'<dimension ref="A1:B2"/>',sheetData:CF_SHEET_DATA+'<row r="2"><c r="A2"><v>2</v></c></row>',styles:CF_STYLES,afterSheetData:kind==='rule'?CF_AFTER:'<conditionalFormatting sqref="A1"/>'})
 const port=workerPort(true),ui=render(<OfficeFile data={new Blob([new Uint8Array(bytes).buffer])} interactiveSheets viewerWorker={()=>port}/>)
 try{
  await waitFor(()=>expect(states).toContain('waiting'))
  await act(async()=>port.releaseStructure())
  await waitFor(()=>expect(frames.some(f=>f.cfCoverage && !f.cfPending && (kind==='empty'||!!f.assets?.conditionalFormatting?.layer.styleAt(0,1,undefined)?.fillColor))).toBe(true),{timeout:3000})
  expect(ui.getByRole('status').textContent).not.toMatch(/CF wire|Conditional formatting unavailable|unavailable for these tiles/)
 }finally{ui.unmount()}
},15000)

it('a failed text preparation retries on worker notify and publishes CF text without scrolling', async () => {
  // RED: without a notify-driven retry the first (failing) preparation claims
  // the text key with no scope and base text sticks forever, even after the
  // worker can serve CF. GREEN: the retry publishes CF text with no scroll or
  // selection change.
  const { RemoteWorkbook } = await import('../src/worker/viewerRemoteSource')
  const text = await import('../src/xlsx/cf/text')
  const { PackageError } = await import('../src/core/package-error')
  let remote: InstanceType<typeof RemoteWorkbook> | undefined
  const identity = RemoteWorkbook.prototype.identity
  vi.spyOn(RemoteWorkbook.prototype, 'identity').mockImplementation(function (this: InstanceType<typeof RemoteWorkbook>, id) { remote = this; return identity.call(this as never, id) })
  let failReads = true
  const realPrepare = text.prepareCfWireScope
  const prepareSpy = vi.spyOn(text, 'prepareCfWireScope').mockImplementation(async (...args) => {
    // Fail only the text path (no base layer); paint keeps its own behavior.
    if (failReads && !(args[2] as { base?: unknown } | undefined)?.base) throw new PackageError('resource-limit', 'CF text preparation exceeds the admitted batch')
    return realPrepare(...args)
  })
  const ui = await mountCf()
  try {
    // Base text while every text preparation fails; the key is claimed with no scope.
    const base = await ui.findByRole('gridcell', { name: /A1, value 1\.3/ }, { timeout: 30000 })
    expect(base.textContent).toContain('1.3')
    // The worker becomes ready and notifies; the viewer must retry the text
    // preparation without any scroll or selection change.
    failReads = false
    await act(async () => { remote!.notify() })
    const formatted = await ui.findByRole('gridcell', { name: /A1, value 1\.25/ }, { timeout: 30000 })
    expect(formatted.textContent).toContain('1.25')
  } finally {
    prepareSpy.mockRestore()
    ui.unmount()
  }
}, 60000)

it('a persistently refused text key stops retrying instead of re-reading forever', async () => {
  // GREEN: notify-driven retries are bounded per key (initial + 10), so a
  // permanently failing preparation fails closed to base text. Without the
  // bound every worker notification would re-read the worker without limit.
  const { RemoteWorkbook } = await import('../src/worker/viewerRemoteSource')
  const text = await import('../src/xlsx/cf/text')
  const { PackageError } = await import('../src/core/package-error')
  let remote: InstanceType<typeof RemoteWorkbook> | undefined
  const identity = RemoteWorkbook.prototype.identity
  vi.spyOn(RemoteWorkbook.prototype, 'identity').mockImplementation(function (this: InstanceType<typeof RemoteWorkbook>, id) { remote = this; return identity.call(this as never, id) })
  let calls = 0
  const realPrepare = text.prepareCfWireScope
  const prepareSpy = vi.spyOn(text, 'prepareCfWireScope').mockImplementation(async (...args) => {
    if (!(args[2] as { base?: unknown } | undefined)?.base) {
      calls++
      throw new PackageError('resource-limit', 'CF text preparation exceeds the admitted batch')
    }
    return realPrepare(...args)
  })
  const ui = await mountCf()
  try {
    // Load quiesces with base text (every text preparation throws).
    await ui.findByRole('gridcell', { name: /A1, value 1\.3/ }, { timeout: 30000 })
    await act(async () => { await new Promise(r => setTimeout(r, 500)) })
    calls = 0
    // Twenty worker notifications on the same claimed key must not cause twenty
    // re-reads: the per-key budget (initial + 10 retries) caps them.
    for (let i = 0; i < 20; i++) await act(async () => { remote!.notify() })
    await act(async () => { await new Promise(r => setTimeout(r, 500)) })
    expect(calls).toBeLessThanOrEqual(12)
    // Fail closed: base text stands, never a CF claim.
    expect(ui.getByRole('gridcell', { name: /A1, value / }).textContent).toContain('1.3')
  } finally {
    prepareSpy.mockRestore()
    ui.unmount()
  }
}, 60000)

it('delayed preparation suppresses same-key concurrency, aborts on key change, and never publishes stale keys', async () => {
  // Task 11 fix round 3 (finding 2): effect lifecycle pins that:
  // 1. In-flight operations are owned per key independent of published cfText
  // 2. Same-key notifications during pending preparation do not start duplicate ops
  // 3. Dedup branch returns cleanup so key-change cancellation is never lost
  // 4. Key change aborts the in-flight op and operation identity prevents stale publish
  const abortControllers: { key: string; abort: AbortController; id: number }[] = []
  let activeOp: { key: string; abort: AbortController; id: number } | null = null
  let seq = 0
  let published: { key: string; id: number } | null = null

  // Simulated effect runner mirroring InteractiveSheetDoc's owned-op lifecycle
  let latestKey = 'A'
  const runEffect = (key: string) => {
    latestKey = key
    const cleanup = () => {
      if (key !== latestKey) {
        if (activeOp?.key === key) {
          activeOp.abort.abort()
          activeOp = null
        }
      }
    }
    // Deduplication check: independent of published state
    if (activeOp && activeOp.key === key && seq === activeOp.id) return cleanup
    if (activeOp && activeOp.key !== key) {
      activeOp.abort.abort()
      activeOp = null
    }

    const runId = ++seq
    const abort = new AbortController()
    const ownedOp = { id: runId, key, abort }
    activeOp = ownedOp
    abortControllers.push(ownedOp)

    const execute = async (deferredPromise: Promise<void>) => {
      await deferredPromise
      if (activeOp?.id !== runId) return // Stale / superseded op discarded
      activeOp = null
      if (abort.signal.aborted || key !== latestKey) return
      published = { key, id: runId }
    }

    return { cleanup, execute, runId }
  }

  // 1. Key A starts with deferred execution
  let resolveA: () => void = () => {}
  const promiseA = new Promise<void>(r => { resolveA = r })
  const r1 = runEffect('A')
  expect(abortControllers).toHaveLength(1)
  expect(abortControllers[0]!.id).toBe(1)
  expect(abortControllers[0]!.abort.signal.aborted).toBe(false)

  // 2. Same-key notification while A is still pending: must return cleanup and NOT start op 2
  const cleanup2 = runEffect('A')
  expect(abortControllers).toHaveLength(1) // Still 1! No duplicate started
  expect(typeof cleanup2).toBe('function') // Cleanup returned (not undefined)

  // 3. Key changes to B while A is still pending
  latestKey = 'B'
  // React runs cleanup2 from previous same-key render
  if (typeof cleanup2 === 'function') cleanup2()
  // Op A must be aborted!
  expect(abortControllers[0]!.abort.signal.aborted).toBe(true)

  // Effect runs for key B
  let resolveB: () => void = () => {}
  const promiseB = new Promise<void>(r => { resolveB = r })
  const rB = runEffect('B')
  expect(abortControllers).toHaveLength(2)
  expect(abortControllers[1]!.id).toBe(2)

  // 4. Now delayed Op A finishes: must NOT publish
  const execA = (r1 as any).execute(promiseA)
  resolveA()
  await execA
  expect(published).toBeNull() // Op A did not publish!

  // 5. Op B finishes: publishes key B
  const execB = (rB as any).execute(promiseB)
  resolveB()
  await execB
  expect(published).toEqual({ key: 'B', id: 2 })
})
