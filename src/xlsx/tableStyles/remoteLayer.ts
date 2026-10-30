/**
 * A region style layer decoded from the worker's bounded descriptors (RemoteRegionStyles for styled
 * tables, RemotePivotDisplay for pivots). Lazy: the main-thread receiver (viewerRemoteRegionStyles.ts)
 * and worker-backed Find (viewerSession.ts) import it dynamically, only for a sheet that has styled
 * tables or pivots, so neither a plain workbook nor the paged worker's own chunks load style code.
 */
import type { XlsxSheet } from '../types'
import { regionLayerFixedBytes, type RemotePivotDisplay, type RemoteRegionStyles } from '../view/contracts'
import type { DxfStyle, TableStyleDef, TableStyleDiagnostic } from './model'
import type { RegionStyleLayer } from './regionLayer'

/** styles.xml dxf ids stop at parseStylesExtras' MAX_DXFS (65,536). */
const MAX_DXF_ID = 65_536

/** The decoded layer's admission and retained charge: the fixed allowance plus, for pivots, each
 * saved layout (lines, item vectors and parent bands; pivots/savedLayout.ts). */
export async function decodedRegionLayerBytes(payload: RemoteRegionStyles | RemotePivotDisplay): Promise<number> {
  let bytes = regionLayerFixedBytes(payload)
  if (bytes && 'pivotTables' in payload && payload.pivotTables.length) {
    const { savedPivotLayoutBytes } = await import('../pivots/savedLayout')
    for (const def of payload.pivotTables) bytes += savedPivotLayoutBytes(def)
  }
  return bytes
}

/** Build the layer the payload describes; `report` receives what building it diagnoses. */
export async function buildRemoteRegionLayer(sheetId: string, payload: RemoteRegionStyles | RemotePivotDisplay, report: (diagnostic: TableStyleDiagnostic) => void): Promise<RegionStyleLayer | undefined> {
  const { preloadTableStyles } = await import('../tableStyleLoader')
  await preloadTableStyles()
  const { buildSheetRegionLayer } = await import('./regionLayer')
  const tables = 'tables' in payload ? payload.tables : []
  const tableStyles = new Map<string, TableStyleDef>(tables.map(item => [item.def.name, item.def]))
  if ('pivotStyles' in payload) for (const def of payload.pivotStyles) tableStyles.set(def.name, def)
  const slicerStyles = new Map()
  // The pivots' `formats` dxfs at their styles.xml ids (a sparse list, as regions.ts indexes it).
  const dxfs: DxfStyle[] = []
  if ('pivotDxfs' in payload) for (const { id, dxf } of payload.pivotDxfs ?? []) if (Number.isSafeInteger(id) && id >= 0 && id < MAX_DXF_ID) dxfs[id] = dxf
  const extras = { dxfs, tableStyles, slicerStyles }
  const sheet = { name: sheetId, tables: tables.map(item => item.table), ...('pivotTables' in payload ? { pivotTables: payload.pivotTables } : {}) } as XlsxSheet
  if (sheet.pivotTables?.length) await (await import('../pivotLoader')).ensurePivotEngine([sheet])
  return buildSheetRegionLayer(sheet, extras, payload.theme ?? (await import('../../drawing/style')).parseThemeContext(), report)
}
