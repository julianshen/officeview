/**
 * Builders for minimal but structurally-valid ODF (OpenDocument) test files.
 * Formatting lives in styles (like real Writer output): the builder mints
 * automatic paragraph/text/table/column styles (P1…, T1…, TT1…, TCW1…)
 * unless a spec names an explicit style defined via the raw-XML escape
 * hatches (extraAutoStyles / extraCommonStyles).
 */

export const ODF_MIMETYPE = 'application/vnd.oasis.opendocument.text'

export interface OdfRunSpec {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  sizePt?: number
  color?: string // RRGGBB
  background?: string // RRGGBB
  /** Explicit text:style-name (must exist in extraAutoStyles); skips T-style minting. */
  spanStyle?: string
  /** Emits text:page-number / text:page-count with cached text. */
  field?: 'PAGE' | 'NUMPAGES'
  /** Raw inline XML fragment (tabs, spaces, breaks); used verbatim. */
  raw?: string
}

export interface OdfParaSpec {
  runs?: Array<string | OdfRunSpec>
  /** Shorthand for a single plain-text run. */
  text?: string
  /** Explicit text:style-name; skips P-style minting for this paragraph. */
  style?: string
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Emits text:h with this outline level instead of text:p. */
  heading?: number
  /** style:master-page-name — a change starts a new section. */
  masterPage?: string
  indentLeftCm?: number
  spacingBeforePt?: number
  spacingAfterPt?: number
  /** Draw frames living inside this paragraph (images). */
  frames?: OdfFrameSpec[]
}

export interface OdfListItem {
  para?: OdfParaSpec
  text?: string
  nested?: OdfListSpec
}

export interface OdfListSpec {
  /** Omitted for a nested list inheriting the surrounding style. */
  styleName?: string
  items: OdfListItem[]
  continueNumbering?: boolean
  startValue?: number
}

export interface OdfCellSpec {
  text?: string
  paras?: OdfParaSpec[]
  gridSpan?: number
  /** table:number-rows-spanned — opens a vertical merge. */
  rowSpan?: number
  /** Emits a covered placeholder cell (span leftover). */
  covered?: boolean
  /** table:number-columns-repeated — stamps identical cells (or placeholders). */
  repeat?: number
  fill?: string
  vAlign?: 'top' | 'middle' | 'bottom'
}

export interface OdfRowSpec {
  cells: OdfCellSpec[]
  heightCm?: number
  exactHeight?: boolean
  /** table:number-rows-repeated — stamps identical rows. */
  repeat?: number
}

export interface OdfTableSpec {
  colWidths: string[]
  rows: OdfRowSpec[]
  /** Rows wrapped in table:table-header-rows (repeat on continuation pages). */
  headerRows?: OdfRowSpec[]
  tableBorders?: string // raw fo:border-* fragment for the table style
  cellBorders?: string // raw fo:border-* fragment applied to every cell style
  relWidths?: boolean // emit style:rel-column-width instead of absolute widths
}

export interface OdfFrameSpec {
  picture: string // Pictures/… path key into pictures
  widthCm: number
  heightCm: number
  anchor?: 'as-char' | 'paragraph' | 'page' | 'char'
  xCm?: number
  yCm?: number
  wrap?: 'none' | 'left' | 'right' | 'parallel' | 'dynamic' | 'run-through'
  behind?: boolean
  zIndex?: number
  hPos?: 'from-left' | 'left' | 'center' | 'right'
  vPos?: 'from-top' | 'top' | 'center' | 'bottom'
}

export interface OdfPageLayoutSpec {
  widthCm: number
  heightCm: number
  marginCm?: number | { top?: number; right?: number; bottom?: number; left?: number }
  landscape?: boolean
}

export interface OdfMasterSpec {
  layout: string
  header?: OdfParaSpec[]
  footer?: OdfParaSpec[]
}

export interface OdfBuildOptions {
  paras?: OdfParaSpec[]
  lists?: OdfListSpec[]
  tables?: OdfTableSpec[]
  /** Raw XML blocks appended to office:text (edge cases: annotations, changes). */
  extraBlocks?: string[]
  pictures?: Record<string, Uint8Array>
  pageLayouts?: Record<string, OdfPageLayoutSpec>
  masters?: Record<string, OdfMasterSpec>
  extraAutoStyles?: string // raw XML inside content.xml automatic-styles
  extraCommonStyles?: string // raw XML inside styles.xml office:styles
}

const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
  + 'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
  + 'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
  + 'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
  + 'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
  + 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
  + 'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
  + 'xmlns:xlink="http://www.w3.org/1999/xlink"'

/** Canned bullet list style (3 levels, •/–/·) for tests. */
export function odfBulletListStyle(name: string): string {
  const level = (n: number, char: string, indent: number): string =>
    `<text:list-level-style-bullet text:level="${n}" text:bullet-char="${char}"><style:list-level-properties fo:margin-left="${indent}cm" fo:text-indent="-0.6cm" fo:min-label-width="0.6cm"/></text:list-level-style-bullet>`
  return `<text:list-style style:name="${name}">${level(1, '•', 1.2)}${level(2, '–', 1.8)}${level(3, '·', 2.4)}</text:list-style>`
}

/** Canned decimal list style (3 levels: 1./a)/i)) for tests. */
export function odfDecimalListStyle(name: string): string {
  return `<text:list-style style:name="${name}">`
    + `<text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties fo:margin-left="1.2cm" fo:text-indent="-0.6cm" fo:min-label-width="0.6cm"/></text:list-level-style-number>`
    + `<text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"><style:list-level-properties fo:margin-left="1.8cm" fo:text-indent="-0.6cm" fo:min-label-width="0.6cm"/></text:list-level-style-number>`
    + `<text:list-level-style-number text:level="3" style:num-format="i" style:num-suffix=")"><style:list-level-properties fo:margin-left="2.4cm" fo:text-indent="-0.6cm" fo:min-label-width="0.6cm"/></text:list-level-style-number>`
    + `</text:list-style>`
}

function esc(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

class Ctx {
  autoStyles: string[] = []
  paraSeq = 0
  textSeq = 0
  tableSeq = 0
  colSeq = 0
  graphicSeq = 0

  runXml(run: string | OdfRunSpec): string {
    const spec: OdfRunSpec = typeof run === 'string' ? { text: run } : run
    if (spec.raw !== undefined) return spec.raw
    if (spec.field === 'PAGE') return `<text:page-number>${esc(spec.text)}</text:page-number>`
    if (spec.field === 'NUMPAGES') return `<text:page-count>${esc(spec.text)}</text:page-count>`
    if (spec.spanStyle) return `<text:span text:style-name="${spec.spanStyle}">${esc(spec.text)}</text:span>`
    if (
      spec.bold === undefined && spec.italic === undefined && spec.underline === undefined
      && spec.strike === undefined && spec.sizePt === undefined && spec.color === undefined
      && spec.background === undefined
    ) {
      return esc(spec.text)
    }
    this.textSeq += 1
    const name = `T${this.textSeq}`
    const props: string[] = []
    if (spec.bold !== undefined) props.push(`fo:font-weight="${spec.bold ? 'bold' : 'normal'}"`)
    if (spec.italic !== undefined) props.push(`fo:font-style="${spec.italic ? 'italic' : 'normal'}"`)
    if (spec.underline !== undefined) props.push(`style:text-underline-style="${spec.underline ? 'solid' : 'none'}" style:text-underline-width="auto" style:text-underline-color="font-color"`)
    if (spec.strike !== undefined) props.push(`style:text-line-through-style="${spec.strike ? 'solid' : 'none'}"`)
    if (spec.sizePt !== undefined) props.push(`fo:font-size="${spec.sizePt}pt"`)
    if (spec.color !== undefined) props.push(`fo:color="#${spec.color}"`)
    if (spec.background !== undefined) props.push(`fo:background-color="#${spec.background}"`)
    this.autoStyles.push(
      `<style:style style:name="${name}" style:family="text"><style:text-properties ${props.join(' ')}/></style:style>`,
    )
    return `<text:span text:style-name="${name}">${esc(spec.text)}</text:span>`
  }

  paraXml(spec: OdfParaSpec): string {
    const runs = (spec.runs ?? (spec.text !== undefined ? [spec.text] : [])).map((r) => this.runXml(r)).join('')
    const frames = (spec.frames ?? []).map((f) => this.frameXml(f)).join('')
    const content = `${runs}${frames}`
    const tag = spec.heading !== undefined ? `text:h text:outline-level="${spec.heading}"` : 'text:p'
    const styleAttr = spec.style ? ` text:style-name="${spec.style}"` : this.mintParaStyle(spec)
    const masterAttr = spec.masterPage ? ` style:master-page-name="${spec.masterPage}"` : ''
    return `<${tag}${styleAttr}${masterAttr}>${content}</${spec.heading !== undefined ? 'text:h' : 'text:p'}>`
  }

  private mintParaStyle(spec: OdfParaSpec): string {
    const props: string[] = []
    if (spec.align && spec.align !== 'left') props.push(`fo:text-align="${spec.align === 'justify' ? 'justify' : spec.align}"`)
    if (spec.indentLeftCm !== undefined) props.push(`fo:margin-left="${spec.indentLeftCm}cm"`)
    if (spec.spacingBeforePt !== undefined) props.push(`fo:margin-top="${spec.spacingBeforePt}pt"`)
    if (spec.spacingAfterPt !== undefined) props.push(`fo:margin-bottom="${spec.spacingAfterPt}pt"`)
    if (props.length === 0) return ''
    this.paraSeq += 1
    const name = `P${this.paraSeq}`
    this.autoStyles.push(
      `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties ${props.join(' ')}/></style:style>`,
    )
    return ` text:style-name="${name}"`
  }

  listXml(list: OdfListSpec): string {
    const attrs = (list.styleName ? `text:style-name="${list.styleName}"` : '')
      + (list.continueNumbering ? ' text:continue-numbering="true"' : '')
      + (list.startValue !== undefined ? ` text:start-value="${list.startValue}"` : '')
    const items = list.items.map((item) => {
      const inner = [
        item.text !== undefined ? this.paraXml({ text: item.text }) : '',
        item.para ? this.paraXml(item.para) : '',
        item.nested ? this.listXml(item.nested) : '',
      ].join('')
      return `<text:list-item>${inner}</text:list-item>`
    }).join('')
    return `<text:list ${attrs}>${items}</text:list>`
  }

  frameXml(frame: OdfFrameSpec): string {
    const anchor = frame.anchor ?? 'as-char'
    const wrap = frame.wrap ?? (anchor === 'as-char' ? 'none' : 'parallel')
    const graphicName = `G${++this.graphicSeq}`
    const graphicProps = [
      `style:wrap="${wrap}"`,
      frame.behind ? 'style:run-through="background"' : '',
      frame.hPos ? `style:horizontal-pos="${frame.hPos}" style:horizontal-rel="paragraph"` : '',
      frame.vPos ? `style:vertical-pos="${frame.vPos}" style:vertical-rel="paragraph"` : '',
    ].filter(Boolean).join(' ')
    this.autoStyles.push(
      `<style:style style:name="${graphicName}" style:family="graphic"><style:graphic-properties ${graphicProps}/></style:style>`,
    )
    const pos = anchor === 'as-char' ? 'svg:x="0cm" svg:y="0cm"' : `svg:x="${frame.xCm ?? 0}cm" svg:y="${frame.yCm ?? 0}cm"`
    const zAttr = frame.zIndex !== undefined ? ` draw:z-index="${frame.zIndex}"` : ''
    return `<draw:frame draw:style-name="${graphicName}" draw:name="Frame${this.graphicSeq}" text:anchor-type="${anchor}" ${pos} svg:width="${frame.widthCm}cm" svg:height="${frame.heightCm}cm"${zAttr}>`
      + `<draw:image xlink:href="${frame.picture}"><text:p/></draw:image></draw:frame>`
  }

  tableXml(table: OdfTableSpec): string {
    this.tableSeq += 1
    const tName = `TT${this.tableSeq}`
    const tableProps = table.tableBorders ? `<style:table-properties ${table.tableBorders}/>` : ''
    this.autoStyles.push(`<style:style style:name="${tName}" style:family="table">${tableProps}</style:style>`)
    const cols = table.colWidths.map((w) => {
      this.colSeq += 1
      const cName = `TCW${this.colSeq}`
      const widthAttr = table.relWidths ? `style:rel-column-width="${w}"` : `style:column-width="${w}"`
      this.autoStyles.push(
        `<style:style style:name="${cName}" style:family="table-column"><style:table-column-properties ${widthAttr}/></style:style>`,
      )
      return `<table:table-column table:style-name="${cName}"/>`
    }).join('')
    const rowXml = (row: OdfRowSpec): string => {
      const cells = row.cells.map((cell) => {
        if (cell.covered) {
          const coverRepeat = cell.repeat && cell.repeat > 1 ? ` table:number-columns-repeated="${cell.repeat}"` : ''
          return `<table:covered-table-cell${coverRepeat}/>`
        }
        const cellProps: string[] = []
        if (cell.fill) cellProps.push(`fo:background-color="#${cell.fill}"`)
        if (cell.vAlign && cell.vAlign !== 'top') {
          cellProps.push(`style:vertical-align="${cell.vAlign}"`)
        }
        if (table.cellBorders) cellProps.push(table.cellBorders)
        const styleAttr = cellProps.length > 0 || cell.fill || cell.vAlign
          ? ` table:style-name="${this.mintCellStyle(cellProps.join(' '))}"`
          : ''
        const spanAttr = cell.gridSpan && cell.gridSpan > 1 ? ` table:number-columns-spanned="${cell.gridSpan}"` : ''
        const rowSpanAttr = cell.rowSpan && cell.rowSpan > 1 ? ` table:number-rows-spanned="${cell.rowSpan}"` : ''
        const repeatAttr = cell.repeat && cell.repeat > 1 ? ` table:number-columns-repeated="${cell.repeat}"` : ''
        const paras = (cell.paras ?? (cell.text !== undefined ? [{ text: cell.text }] : [{}]))
          .map((p) => this.paraXml(p)).join('')
        return `<table:table-cell${styleAttr}${spanAttr}${rowSpanAttr}${repeatAttr}>${paras}</table:table-cell>`
      }).join('')
      const heightAttr = row.heightCm !== undefined
        ? ` table:style-name="${this.mintRowStyle(row.heightCm, row.exactHeight)}"`
        : ''
      const rowRepeatAttr = row.repeat && row.repeat > 1 ? ` table:number-rows-repeated="${row.repeat}"` : ''
      return `<table:table-row${heightAttr}${rowRepeatAttr}>${cells}</table:table-row>`
    }
    const rows = table.rows.map(rowXml).join('')
    const headerRows = (table.headerRows ?? []).map(rowXml).join('')
    const headerXml = headerRows ? `<table:table-header-rows>${headerRows}</table:table-header-rows>` : ''
    return `<table:table table:name="Table${this.tableSeq}" table:style-name="${tName}">${cols}${headerXml}${rows}</table:table>`
  }

  private mintCellStyle(props: string): string {
    this.colSeq += 1
    const name = `TC${this.colSeq}`
    this.autoStyles.push(
      `<style:style style:name="${name}" style:family="table-cell"><style:table-cell-properties ${props}/></style:style>`,
    )
    return name
  }

  private mintRowStyle(heightCm: number, exact?: boolean): string {
    this.colSeq += 1
    const name = `TR${this.colSeq}`
    this.autoStyles.push(
      `<style:style style:name="${name}" style:family="table-row"><style:table-row-properties style:row-height="${heightCm}cm" style:use-optimal-row-height="${exact ? 'false' : 'true'}"/></style:style>`,
    )
    return name
  }
}

/** Build a minimal .odt buffer. */
export async function buildOdt(opts: OdfBuildOptions = {}): Promise<Uint8Array> {
  const { default: JSZip } = await import('jszip')
  const ctx = new Ctx()
  // render content first so style minting completes before serialization
  const paraXml = (opts.paras ?? []).map((p) => ctx.paraXml(p)).join('')
  const listXml = (opts.lists ?? []).map((l) => ctx.listXml(l)).join('')
  const tableXml = (opts.tables ?? []).map((t) => ctx.tableXml(t)).join('')
  const extra = (opts.extraBlocks ?? []).join('')
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>`
    + `<office:document-content ${NS} office:version="1.3">`
    + `<office:font-face-decls>`
    + `<style:font-face style:name="Liberation Serif" svg:font-family="Liberation Serif"/>`
    + `<style:font-face style:name="Arial" svg:font-family="Arial"/>`
    + `</office:font-face-decls>`
    + `<office:automatic-styles>${ctx.autoStyles.join('')}${opts.extraAutoStyles ?? ''}</office:automatic-styles>`
    + `<office:body><office:text>`
    + `${paraXml}${listXml}${tableXml}${extra}`
    + `</office:text></office:body>`
    + `</office:document-content>`
  const layouts = Object.entries(opts.pageLayouts ?? {}).map(([name, spec]) => {
    const m = typeof spec.marginCm === 'number'
      ? { top: spec.marginCm, right: spec.marginCm, bottom: spec.marginCm, left: spec.marginCm }
      : { top: 2, right: 2, bottom: 2, left: 2, ...(spec.marginCm ?? {}) }
    return `<style:page-layout style:name="${name}">`
      + `<style:page-layout-properties fo:page-width="${spec.widthCm}cm" fo:page-height="${spec.heightCm}cm" `
      + `style:print-orientation="${spec.landscape ? 'landscape' : 'portrait'}" `
      + `fo:margin-top="${m.top}cm" fo:margin-right="${m.right}cm" fo:margin-bottom="${m.bottom}cm" fo:margin-left="${m.left}cm"/>`
      + `</style:page-layout>`
  }).join('')
  const masters = Object.entries(opts.masters ?? {}).map(([name, spec]) => {
    const sub = new Ctx()
    const header = (spec.header ?? []).map((p) => sub.paraXml(p)).join('')
    const footer = (spec.footer ?? []).map((p) => sub.paraXml(p)).join('')
    return `<style:master-page style:name="${name}" style:page-layout-name="${spec.layout}">`
      + (header ? `<style:header>${header}</style:header>` : '')
      + (footer ? `<style:footer>${footer}</style:footer>` : '')
      + `</style:master-page>`
  }).join('')
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>`
    + `<office:document-styles ${NS} office:version="1.3">`
    + `<office:font-face-decls>`
    + `<style:font-face style:name="Liberation Serif" svg:font-family="Liberation Serif"/>`
    + `</office:font-face-decls>`
    + `<office:styles>`
    + `<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.25cm"/><style:text-properties fo:font-family="Liberation Serif" fo:font-size="12pt"/></style:default-style>`
    + `${opts.extraCommonStyles ?? ''}`
    + `</office:styles>`
    + `<office:automatic-styles>${layouts}</office:automatic-styles>`
    + `<office:master-styles>${masters || '<style:master-page style:name="Standard" style:page-layout-name="__default__"/>'}</office:master-styles>`
    + `</office:document-styles>`
  const pictures = opts.pictures ?? {}
  const manifestEntries = [
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>`,
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>`,
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>`,
    ...Object.keys(pictures).map((p) => {
      const ext = p.split('.').pop()?.toLowerCase()
      const media = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/png'
      return `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${media}"/>`
    }),
  ].join('')
  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>`
    + `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">${manifestEntries}</manifest:manifest>`
  const zip = new JSZip()
  zip.file('mimetype', ODF_MIMETYPE, { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', manifest)
  zip.file('content.xml', content)
  zip.file('styles.xml', styles)
  for (const [path, data] of Object.entries(pictures)) zip.file(path, data)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}
