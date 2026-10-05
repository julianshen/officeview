/** Print a machine-readable drawing inventory for each supplied Office file. */
import { readFileSync } from 'node:fs'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { drawingReport } from '../src/drawing/coverage'
import { getPaintables } from '../src/render/paint'

async function reportOne(path: string) {
  try {
    const document = await loadOfficeFile(readFileSync(path))
    let format: 'docx' | 'pptx' | 'xlsx'
    let unit: 'page' | 'slide' | 'sheet'
    let unitCount: number
    if ('slides' in document) { format = 'pptx'; unit = 'slide'; unitCount = document.slides.length }
    else if ('sheets' in document) { format = 'xlsx'; unit = 'sheet'; unitCount = document.sheets.length }
    else if ('sections' in document) {
      format = 'docx'; unit = 'page'
      const pages = await getPaintables(document)
      try { unitCount = pages.length } finally { pages.dispose() }
    } else throw new Error('Drawing coverage supports DOCX, PPTX, and XLSX')
    return { path, format, unit, unitCount, ...drawingReport(document) }
  } catch (error) {
    return { path, error: error instanceof Error ? error.message : String(error) }
  }
}

const paths = process.argv.slice(2)
const files = await Promise.all(paths.map(reportOne))
const report = {
  schemaVersion: 1,
  countDefinitions: {
    entryUnit: 'Zero-based slide or sheet index when known. DOCX entries use source part and tree path because flow drawings do not have a fixed page before layout',
    representation: 'Selected markup compatibility branch: native, Choice, or Fallback',
    selectedRepresentation: 'Payload used by the static renderer; cached-diagram and raster-fallback do not claim uncached diagram, math, 3D, or ChartEx algorithms',
    extentUnits: 'Requested and retained viewport extents are in CSS pixels; depth and node limits are counts',
    originalObjects: 'Original top-level Word placements, slide objects, or worksheet anchors/containers; one per source object, regardless of selected descendants or diagnostic facets',
    selectedDescendants: 'Selected nested source objects below original containers, including group children',
    diagnosticEntries: 'Failures or resource limits without another source-object placement entry',
  },
  files,
}
try { process.stdout.write(`${JSON.stringify(report)}\n`) }
catch (error) { process.stderr.write(`Drawing report serialization failed: ${error instanceof Error ? error.message : String(error)}\n`); process.exitCode = 1 }
if (!paths.length || files.some(file => 'error' in file)) process.exitCode = 1
