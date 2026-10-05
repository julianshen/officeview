/**
 * Author-generated CI fixtures for SVG/raster picture selection (Stage 3 Task 4).
 *
 * Compact, self-contained packages built with JSZip — no proprietary inputs or
 * binary dumps. The SVG and PNG candidates for each icon deliberately differ in
 * BOTH colour and geometry so a pixel test can prove which representation was
 * selected (the protected native icons share art between SVG and PNG).
 */
import JSZip from 'jszip'

export const SVG_NS = 'http://www.w3.org/2000/svg'
export const ASVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
export const XLINK_NS = 'http://www.w3.org/1999/xlink'
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const SVG_BLIP_GUID = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}'
/** Discard port: never a real endpoint; a canary for accidental network use. */
export const CANARY_ORIGIN = 'http://127.0.0.1:9'

/**
 * Portable base64 decode: browsers have `atob` but no `Buffer`, and this runs at
 * module initialization. `Buffer` is only referenced as a guarded fallback.
 */
const b64 = (value: string): Uint8Array => {
  if (typeof atob === 'function') {
    const binary = atob(value)
    const out = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i)
    return out
  }
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(value, 'base64'))
  throw new Error('no base64 decoder available')
}
export const svgBytes = (text: string): Uint8Array => new TextEncoder().encode(text)

/** 32x32 author-rendered PNG fallbacks (distinct colour AND geometry per icon). */
export const PNG = {
  redX: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAA10lEQVRYhe2Xyw6CQAxFr4aIf+xSln6PC+PHmdQNJBNgxtsHlgVNWJH2nj4YOsBhezcBugzfKcAgwFuAq8H3IsBTgIdHXMZHBVGIT/46iJm4CmJFXAdREacgGuIchACdAK9GABnf9yu+PenbHkwii0UlLD5hEOHihpKaWhZZidjMAyH84g6IOPECgpkJdc/POgZ8yJgnRVxKOa8FqUOY+hmSQ7fNQURmvs1RbAkYBiE7+R3frFkQlbj/7kEdwruSceIVCO9SqhOfQXjXcpt4ESjvYnLYP+wLtbyhLda1yM8AAAAASUVORK5CYII='),
  cyanSquare: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAAN0lEQVRYhe3OoREAMAgEQUj/PZMCEoNgMHvyzW+EtFx+16qaecvn74wcNQIAAAAAAAAAAACQ1rv9ugMs3LZpdwAAAABJRU5ErkJggg=='),
  magentaTriangle: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAAhklEQVRYhe3TQQrAMAhEUY/mzXO06aaLEghJdFQKDmRd34eK9Hp/HgQKgVYeMCAYVR9XCPA+rThgfA7IrTDp8ytM+twKC31ehYU+p8JGH19ho4+tcKiPq3Coj6lwqedXuNRzKxj1vApGPaeCU++v4NT7KpD09gokva0CWX9fgazn/BG9XtQezEb/HQfdQz8AAAAASUVORK5CYII='),
  navyCircle: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAABhklEQVRYhe3Wv2tTUXTH8U+qqF0qZhKbxU10E6UdHAqdbfSPcJciilDoHYRChwj+Df0Dku5CxyIFnXQs9IejpgQpQjEOSbDDvXm95j265Lue88753vfuefcyZcolU8tLD3U8wwruoTEMHOEbtW2utXn7s2SB1iy9VfqvMVeQfIJN5t6zelqCQLiDNh5fwPQ8X9AkHEwgEBrYxXxm8xFHXF1k7TiVMJN+tjWLzgTNocFZm3DjPwR6r/BwguYjHuFlKpj4BKGOfcUb7qKccP1ubDoSb6D2vMTmcJPfzVggIdB/WmLzESsZAh5UIHA/R+B2BQLRaRozBaXzJ0fgewUC0Zopga8VCERrpsZwuwKBToZAv2NwqpVFF9FFXYnn75yyNIPlkgTWCR9jgXFT0MJeCc0/4UMqWHAcv5vnbNe/m08uh1gkJKeq4D+wdowFg1Xk8hlPxjUnuQfOs9NjaQu/DG5FybN9SBfr1F/w5kdR9cxL6cat4anWFL+UdgYTFLp5dadMuUT+ApjATKMA+d/9AAAAAElFTkSuQmCC'),
  greenSolid: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAAMElEQVRYhe3OMQEAIAzAsIJ/afMEMvakBprT9Frsbs4BAAAAAAAAAAAAAAAAAACqPvmuAgslLcp2AAAAAElFTkSuQmCC'),
  blueSolid: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAAMElEQVRYhe3OMQEAIAzAsIJ/afMEMvakBppT81rsbs4BAAAAAAAAAAAAAAAAAACqPvjiAgui5PL6AAAAAElFTkSuQmCC'),
  purpleSolid: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAAL0lEQVRYhe3OMQEAIAzAsILyOQcZe1IDzZnmtdjdnAMAAAAAAAAAAAAAAAAAAFR9QAICP/aMf5kAAAAASUVORK5CYII='),
  orangeSolid: b64('iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0QA/wD/AP+gvaeTAAAAL0lEQVRYhe3OMQEAIAzAsIJynA8Ze1IDzZnXtNjdnAMAAAAAAAAAAAAAAAAAAFR9QFsCvh3NQPMAAAAASUVORK5CYII='),
}

const svgDoc = (inner: string): string =>
  `<svg xmlns="${SVG_NS}" xmlns:xlink="${XLINK_NS}" width="32" height="32" viewBox="0 0 32 32">${inner}</svg>`

export interface IconFixture {
  name: string
  svg: Uint8Array
  png: Uint8Array
  svgColor: string
  pngColor: string
}

/** Four icons whose SVG and PNG candidates differ in colour and shape. */
export const ICONS: IconFixture[] = [
  { name: 'icon1', svg: svgBytes(svgDoc('<path d="M6 17 L13 24 L26 8" fill="none" stroke="#5B9BD5" stroke-width="5"/>')), png: PNG.redX, svgColor: '#5B9BD5', pngColor: '#FF0000' },
  { name: 'icon2', svg: svgBytes(svgDoc('<path d="M16 3 L20 12 L30 12 L22 18 L25 28 L16 22 L7 28 L10 18 L2 12 L12 12 Z" fill="#FFC000"/>')), png: PNG.cyanSquare, svgColor: '#FFC000', pngColor: '#00FFFF' },
  { name: 'icon3', svg: svgBytes(svgDoc('<path d="M4 16 H22 M22 16 L14 8 M22 16 L14 24" fill="none" stroke="#70AD47" stroke-width="5"/>')), png: PNG.magentaTriangle, svgColor: '#70AD47', pngColor: '#FF00FF' },
  { name: 'icon4', svg: svgBytes(svgDoc('<path d="M8 8 L24 24 M24 8 L8 24" fill="none" stroke="#ED7D31" stroke-width="5"/>')), png: PNG.navyCircle, svgColor: '#ED7D31', pngColor: '#000080' },
]

/** Full SVG documents for the malicious matrix (25 reject + 2 keep). */
export const MALICIOUS_MATRIX: Array<{ name: string; verdict: 'reject' | 'keep'; svg: Uint8Array }> = [
  { name: 'm01_image_href', verdict: 'reject', svg: svgBytes(svgDoc(`<image href="${CANARY_ORIGIN}/c.png"/>`)) },
  { name: 'm02_image_xlink', verdict: 'reject', svg: svgBytes(svgDoc(`<image xlink:href="${CANARY_ORIGIN}/c.png"/>`)) },
  { name: 'm03_image_prefix', verdict: 'reject', svg: svgBytes(svgDoc(`<image foo:href="${CANARY_ORIGIN}/c.png" xmlns:foo="${XLINK_NS}"/>`)) },
  { name: 'm04_image_file', verdict: 'reject', svg: svgBytes(svgDoc('<image href="file:///tmp/canary.txt"/>')) },
  { name: 'm05_use_href', verdict: 'reject', svg: svgBytes(svgDoc(`<use href="${CANARY_ORIGIN}/x.svg#a"/>`)) },
  { name: 'm06_feImage', verdict: 'reject', svg: svgBytes(svgDoc(`<filter id="f"><feImage href="${CANARY_ORIGIN}/c.png"/></filter>`)) },
  { name: 'm07_gradient_href', verdict: 'reject', svg: svgBytes(svgDoc(`<linearGradient href="${CANARY_ORIGIN}/g.svg"/>`)) },
  { name: 'm08_pattern_href', verdict: 'reject', svg: svgBytes(svgDoc(`<pattern href="${CANARY_ORIGIN}/p.svg"/>`)) },
  { name: 'm09_mpath', verdict: 'reject', svg: svgBytes(svgDoc(`<mpath href="${CANARY_ORIGIN}/m.svg"/>`)) },
  { name: 'm10_cursor', verdict: 'reject', svg: svgBytes(svgDoc(`<cursor href="${CANARY_ORIGIN}/c.svg"/>`)) },
  { name: 'm11_css_import', verdict: 'reject', svg: svgBytes(svgDoc(`<style>@import url("${CANARY_ORIGIN}/a.css");</style>`)) },
  { name: 'm12_css_url', verdict: 'reject', svg: svgBytes(svgDoc(`<style>rect{fill:url(${CANARY_ORIGIN}/a.svg#g)}</style>`)) },
  { name: 'm13_fontface', verdict: 'reject', svg: svgBytes(svgDoc(`<style>@font-face{src:url("${CANARY_ORIGIN}/f.woff")}</style>`)) },
  { name: 'm14_css_obfusc', verdict: 'reject', svg: svgBytes(svgDoc(`<style>rect{fill:url( /*c*/ "${CANARY_ORIGIN}/a.svg" )}</style>`)) },
  { name: 'm15_css_case', verdict: 'reject', svg: svgBytes(svgDoc(`<style>RECT{FILL:URL("${CANARY_ORIGIN}/a.svg")}</style>`)) },
  { name: 'm16_xml_stylesheet', verdict: 'reject', svg: svgBytes(`<?xml-stylesheet href="${CANARY_ORIGIN}/a.css"?><svg xmlns="${SVG_NS}"/>`) },
  { name: 'm17_entity_ext', verdict: 'reject', svg: svgBytes(`<!DOCTYPE svg [<!ENTITY x SYSTEM "file:///tmp/canary.txt">]><svg xmlns="${SVG_NS}">&x;</svg>`) },
  { name: 'm18_entity_param', verdict: 'reject', svg: svgBytes(`<!DOCTYPE svg [<!ENTITY % p SYSTEM "${CANARY_ORIGIN}/e">%p;]><svg xmlns="${SVG_NS}"/>`) },
  { name: 'm19_entity_expand', verdict: 'reject', svg: svgBytes(`<!DOCTYPE svg [<!ENTITY a "x"><!ENTITY b "&a;&a;&a;">]><svg xmlns="${SVG_NS}">&b;</svg>`) },
  { name: 'm20_script', verdict: 'reject', svg: svgBytes(svgDoc(`<script>fetch("${CANARY_ORIGIN}/c")</script>`)) },
  { name: 'm21_event', verdict: 'reject', svg: svgBytes(svgDoc(`<rect onload="fetch('${CANARY_ORIGIN}/c')"/>`)) },
  { name: 'm22_foreignObject', verdict: 'reject', svg: svgBytes(svgDoc(`<foreignObject><iframe src="${CANARY_ORIGIN}/c"/></foreignObject>`)) },
  { name: 'm23_data_svg', verdict: 'reject', svg: svgBytes(svgDoc('<image href="data:image/svg+xml;base64,PHN2Zz48c2NyaXB0Pg=="/>')) },
  { name: 'm24_data_html', verdict: 'reject', svg: svgBytes(svgDoc(`<image href="data:text/html,<script>fetch('${CANARY_ORIGIN}/c')</script>"/>`)) },
  { name: 'm25_svgz', verdict: 'reject', svg: b64('H4sIAAAAAAAAE3XMMQ7CMAxA0atYZq+NDQwoydCr0JBEChQ1UdzjI8Tc7S/vuzYS7K/6bh5z7587kZlNptO6JRJmpjYSgpWlZ48qCDmWlPu/R4k2r7tHBgYVUMHgtvjoR+BZavV4Ooterjek4H778AXLFKIAhgAAAA==') },
  { name: 'p01_use_local', verdict: 'keep', svg: svgBytes(svgDoc('<defs><g id="local"><rect width="32" height="32" fill="#5B9BD5"/></g></defs><use href="#local"/>')) },
  { name: 'p02_paint_server', verdict: 'keep', svg: svgBytes(svgDoc('<defs><linearGradient id="grad"><stop offset="0" stop-color="#5B9BD5"/><stop offset="1" stop-color="#ED7D31"/></linearGradient></defs><rect width="32" height="32" fill="url(#grad)"/>')) },
]

/** Deterministic gzip (SVGZ) bytes embedded as base64 (browser-bundleable). */

// ---------------------------------------------------------------------------
// OOXML package builders
// ---------------------------------------------------------------------------

interface MediaSpec { name: string; data: Uint8Array; external?: boolean; target?: string }
export interface PictureSpec {
  id: number
  png?: MediaSpec
  /** svgBlip candidate. `link` emits r:link instead of r:embed. */
  svg?: MediaSpec & { link?: boolean }
  extUri?: string
  /** svgBlip namespace prefix (namespace is unchanged; prefix is arbitrary). */
  svgPrefix?: string
  widthPx?: number
  heightPx?: number
}
export interface ParaSpec { text?: string; pic?: PictureSpec }

const xmlEscape = (value: string): string => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const emu = (px: number): number => Math.round(px * 9525)

function mediaFiles(specs: PictureSpec[]): Map<string, MediaSpec> {
  const out = new Map<string, MediaSpec>()
  for (const spec of specs) for (const media of [spec.png, spec.svg]) if (media && !media.external) out.set(media.name, media)
  return out
}
function pictureSpecs(paras: ParaSpec[]): PictureSpec[] {
  return paras.flatMap(p => p.pic ? [p.pic] : [])
}

// ---- DOCX -----------------------------------------------------------------

const DOCX_CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`
const DOCX_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`
const DOCX_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>`

function docxDrawing(spec: PictureSpec): string {
  const width = emu(spec.widthPx ?? 48)
  const height = emu(spec.heightPx ?? 48)
  const prefix = spec.svgPrefix ?? 'asvg'
  const extUri = spec.extUri ?? SVG_BLIP_GUID
  const pngRid = spec.png ? `png${spec.id}` : undefined
  const svgRid = spec.svg ? `svg${spec.id}` : undefined
  const svgAttr = spec.svg ? (spec.svg.link ? `r:link="${svgRid}"` : `r:embed="${svgRid}"`) : undefined
  const ext = spec.svg
    ? `<a:extLst><a:ext uri="${extUri}"><${prefix}:svgBlip xmlns:${prefix}="${ASVG_NS}" ${svgAttr}/></a:ext></a:extLst>`
    : ''
  return `<w:r><w:drawing><wp:inline><wp:extent cx="${width}" cy="${height}"/><wp:docPr id="${spec.id}" name="Picture ${spec.id}"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${spec.id}" name="Picture ${spec.id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip${pngRid ? ` r:embed="${pngRid}"` : ''}>${ext}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
}
function docxPara(para: ParaSpec): string {
  if (para.pic) return `<w:p>${docxDrawing(para.pic)}</w:p>`
  return `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(para.text ?? '')}</w:t></w:r></w:p>`
}
function docxPart(root: string, paras: ParaSpec[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="${ASVG_NS}">${paras.map(docxPara).join('')}</w:${root}>`
}
function docxRels(specs: PictureSpec[], extra = ''): string {
  const rows: string[] = []
  for (const spec of specs) {
    if (spec.png) rows.push(`<Relationship Id="png${spec.id}" Type="${R_NS}/image" Target="${xmlEscape(spec.png.target ?? `media/${spec.png.name}`)}"${spec.png.external ? ' TargetMode="External"' : ''}/>`)
    if (spec.svg) rows.push(`<Relationship Id="svg${spec.id}" Type="${R_NS}/image" Target="${xmlEscape(spec.svg.target ?? `media/${spec.svg.name}`)}"${spec.svg.external ? ' TargetMode="External"' : ''}/>`)
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rows.join('')}${extra}</Relationships>`
}
async function writeMedia(zip: JSZip, prefix: string, specs: PictureSpec[]): Promise<void> {
  for (const media of mediaFiles(specs).values()) zip.file(`${prefix}${media.name}`, media.data)
}

/** DOCX with body paragraphs, optional header/footer parts and a narrow table. */
export async function buildSvgDocx(
  paras: ParaSpec[],
  opts: { header?: ParaSpec[]; footer?: ParaSpec[]; narrow?: ParaSpec[] } = {},
): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', DOCX_CT)
  zip.file('_rels/.rels', DOCX_ROOT_RELS)
  zip.file('word/styles.xml', DOCX_STYLES)
  const all = [...paras, ...(opts.header ?? []), ...(opts.footer ?? []), ...(opts.narrow ?? [])]
  const specs = pictureSpecs(all)
  let refs = '', extra = ''
  const relParts: Array<[string, string]> = []
  if (opts.header) {
    refs += '<w:headerReference w:type="default" r:id="hdr"/>'
    extra += `<Relationship Id="hdr" Type="${R_NS}/header" Target="header1.xml"/>`
    zip.file('word/header1.xml', docxPart('hdr', opts.header))
    relParts.push(['word/_rels/header1.xml.rels', docxRels(pictureSpecs(opts.header))])
  }
  if (opts.footer) {
    refs += '<w:footerReference w:type="default" r:id="ftr"/>'
    extra += `<Relationship Id="ftr" Type="${R_NS}/footer" Target="footer1.xml"/>`
    zip.file('word/footer1.xml', docxPart('ftr', opts.footer))
    relParts.push(['word/_rels/footer1.xml.rels', docxRels(pictureSpecs(opts.footer))])
  }
  relParts.unshift(['word/_rels/document.xml.rels', docxRels(pictureSpecs([...paras, ...(opts.narrow ?? [])]), extra)])
  const body = opts.narrow
    ? paras.map(docxPara).join('') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1995"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1995" w:type="dxa"/></w:tcPr>${opts.narrow.map(docxPara).join('')}</w:tc></w:tr></w:tbl>`
    : paras.map(docxPara).join('')
  zip.file('word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="${ASVG_NS}"><w:body>${body}<w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
  for (const [path, xml] of relParts) zip.file(path, xml)
  await writeMedia(zip, 'word/media/', specs)
  return zip.generateAsync({ type: 'uint8array' })
}

/** Four interleaved icons between Before/After plus a narrow three-icon wrap. */
export type InlineRun = { text: string } | { pic: PictureSpec }
const inlinePics = (runs: InlineRun[]): PictureSpec[] => runs.flatMap(r => 'pic' in r ? [r.pic] : [])
const inlineRunXml = (run: InlineRun): string => 'pic' in run ? docxDrawing(run.pic) : `<w:r><w:t xml:space="preserve">${xmlEscape(run.text)}</w:t></w:r>`
const inlinePara = (runs: InlineRun[], align?: string): string => `<w:p>${align ? `<w:pPr><w:jc w:val="${align}"/></w:pPr>` : ''}${runs.map(inlineRunXml).join('')}</w:p>`

/**
 * DOCX whose body is ONE paragraph of interleaved text and image runs, plus an
 * optional narrow single-paragraph table cell that must wrap on available width.
 */
export async function buildInlineDocx(runs: InlineRun[], opts: { align?: 'left' | 'center' | 'right' | 'both'; narrow?: { widthTwips?: number; runs: InlineRun[] } } = {}): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', DOCX_CT)
  zip.file('_rels/.rels', DOCX_ROOT_RELS)
  zip.file('word/styles.xml', DOCX_STYLES)
  const width = opts.narrow?.widthTwips ?? 2000
  const body = inlinePara(runs, opts.align) + (opts.narrow
    ? `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="${width}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${inlinePara(opts.narrow.runs)}</w:tc></w:tr></w:tbl>`
    : '')
  const specs = [...inlinePics(runs), ...(opts.narrow ? inlinePics(opts.narrow.runs) : [])]
  zip.file('word/_rels/document.xml.rels', docxRels(specs))
  zip.file('word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="${ASVG_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
  await writeMedia(zip, 'word/media/', specs)
  return zip.generateAsync({ type: 'uint8array' })
}

/** One paragraph: Before + 4 inline icons + After (image widths must advance the text pen). */
export async function distinctDocx(): Promise<Uint8Array> {
  const runs: InlineRun[] = [{ text: 'Before' }]
  for (let i = 0; i < ICONS.length; i++) runs.push({ pic: { id: i + 1, png: { name: `icon${i + 1}.png`, data: ICONS[i].png }, svg: { name: `icon${i + 1}.svg`, data: ICONS[i].svg } } })
  runs.push({ text: 'After' })
  return buildInlineDocx(runs, { narrow: { widthTwips: 2000, runs: narrowRuns() } })
}

function narrowRuns(): InlineRun[] {
  return [0, 1, 2].map(i => ({ pic: { id: 11 + i, png: { name: `icon${i + 1}.png`, data: ICONS[i].png }, svg: { name: `icon${i + 1}.svg`, data: ICONS[i].svg } } }))
}

/** One paragraph of three 48px icons in a ~133px (2000 twip) cell. */
export async function narrowDocx(): Promise<Uint8Array> {
  return buildInlineDocx([{ text: 'Narrow wrap:' }], { narrow: { widthTwips: 2000, runs: narrowRuns() } })
}

/** Justified one paragraph with a tab and an inline icon between text. */
export async function justifyInlineDocx(): Promise<Uint8Array> {
  return buildInlineDocx([{ text: 'Left\t' }, { pic: { id: 1, png: { name: 'icon1.png', data: ICONS[0].png }, svg: { name: 'icon1.svg', data: ICONS[0].svg } } }, { text: ' Right side' }], { align: 'both' })
}

/** Narrow one paragraph whose text wraps after an inline icon. */
export async function wrappedInlineDocx(): Promise<Uint8Array> {
  return buildInlineDocx([{ text: 'Before' }], { narrow: { widthTwips: 2000, runs: [{ text: 'Wrap ' }, { pic: { id: 21, png: { name: 'icon1.png', data: ICONS[0].png }, svg: { name: 'icon1.svg', data: ICONS[0].svg } } }, { text: ' tail words that must wrap onto following lines' }] } })
}

export async function fallbackDocx(): Promise<Uint8Array> {
  return buildSvgDocx([
    { pic: { id: 1, png: { name: 'missing.png', data: PNG.greenSolid }, svg: { name: 'missing.svg', data: svgBytes(svgDoc('')), target: 'media/absent.svg' } } },
    { pic: { id: 2, png: { name: 'invalid.png', data: PNG.blueSolid }, svg: { name: 'invalid.svg', data: svgBytes('not an svg at all') } } },
    { pic: { id: 3, png: { name: 'external.png', data: PNG.orangeSolid }, svg: { name: 'external.svg', data: svgBytes(svgDoc('')), external: true, target: `${CANARY_ORIGIN}/x.svg` } } },
    { pic: { id: 4, png: { name: 'raster.png', data: PNG.purpleSolid } } },
  ])
}

export async function rasterOnlyDocx(): Promise<Uint8Array> {
  return buildSvgDocx([{ text: 'Before' }, { pic: { id: 1, png: { name: 'raster.png', data: PNG.purpleSolid } } }, { text: 'After' }])
}

/** Bare a:blip referencing an SVG part directly (no svgBlip extension). */
export async function primarySvgDocx(svg: Uint8Array = svgBytes(svgDoc('<rect width="32" height="32" fill="#5B9BD5"/>'))): Promise<Uint8Array> {
  return buildSvgDocx([{ pic: { id: 1, png: { name: 'direct.svg', data: svg } } }])
}

/** Same PNG referenced by two pictures with DIFFERENT SVGs, plus an identical pair. */
export async function aliasDocx(): Promise<Uint8Array> {
  return buildSvgDocx([
    { pic: { id: 1, png: { name: 'shared.png', data: PNG.redX }, svg: { name: 'a.svg', data: svgBytes(svgDoc('<rect width="32" height="32" fill="#5B9BD5"/>')) } } },
    { pic: { id: 2, png: { name: 'shared.png', data: PNG.redX }, svg: { name: 'b.svg', data: svgBytes(svgDoc('<rect width="32" height="32" fill="#70AD47"/>')) } } },
    { pic: { id: 3, png: { name: 'same.png', data: PNG.cyanSquare }, svg: { name: 'c.svg', data: svgBytes(svgDoc('<rect width="32" height="32" fill="#FFC000"/>')) } } },
    { pic: { id: 4, png: { name: 'same.png', data: PNG.cyanSquare }, svg: { name: 'c.svg', data: svgBytes(svgDoc('<rect width="32" height="32" fill="#FFC000"/>')) } } },
  ])
}

/** Body picture plus repeated header/footer pictures with local relationships. */
export async function headerDocx(): Promise<Uint8Array> {
  return buildSvgDocx(
    [{ text: 'Body' }, { pic: { id: 1, png: { name: 'body.png', data: PNG.purpleSolid }, svg: { name: 'body.svg', data: ICONS[0].svg } } }],
    { header: [{ pic: { id: 2, png: { name: 'head.png', data: PNG.greenSolid }, svg: { name: 'head.svg', data: ICONS[1].svg } } }], footer: [{ pic: { id: 3, png: { name: 'foot.png', data: PNG.blueSolid }, svg: { name: 'foot.svg', data: ICONS[2].svg } } }] },
  )
}

/** Malicious matrix as a picture-per-row DOCX: one SVG part per row plus a PNG fallback. */
export async function maliciousDocx(): Promise<Uint8Array> {
  const paras: ParaSpec[] = MALICIOUS_MATRIX.map((row, i) => ({
    pic: { id: i + 1, png: { name: `m${i + 1}.png`, data: PNG.greenSolid }, svg: { name: `m${i + 1}.svg`, data: row.svg } },
  }))
  return buildSvgDocx(paras)
}

/** A textbox whose nested picture carries an svgBlip pair (Task 3 nested flow + SVG). */
export async function nestedTextboxSvgDocx(): Promise<Uint8Array> {
  const zip = new JSZip()
  const inlinePic = (id: string): string =>
    `<w:drawing><wp:inline><wp:extent cx="457200" cy="457200"/><wp:docPr id="${id}" name="nested${id}"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="nested${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="png1"><a:extLst><a:ext uri="${SVG_BLIP_GUID}"><asvg:svgBlip r:embed="svg1"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
  const textbox =
    `<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="10" name="box"/>` +
    `<a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r>${inlinePic('11')}</w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
  zip.file('[Content_Types].xml', DOCX_CT)
  zip.file('_rels/.rels', DOCX_ROOT_RELS)
  zip.file('word/styles.xml', DOCX_STYLES)
  zip.file('word/media/nested.png', PNG.redX)
  zip.file('word/media/nested.svg', ICONS[0].svg)
  zip.file('word/_rels/document.xml.rels', docxRels([{ id: 1, png: { name: 'nested.png', data: PNG.redX }, svg: { name: 'nested.svg', data: ICONS[0].svg } }]))
  zip.file('word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="${ASVG_NS}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>` +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r>${textbox}</w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
  return zip.generateAsync({ type: 'uint8array' })
}

// ---- PPTX -----------------------------------------------------------------

const PPTX_CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`
const PPTX_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
const PPTX_PRESENTATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${R_NS}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`

function pptxPic(spec: PictureSpec, index = 0): string {
  const width = emu(spec.widthPx ?? 48)
  const height = emu(spec.heightPx ?? 48)
  const offsetX = emu(index * ((spec.widthPx ?? 48) + 12))
  const prefix = spec.svgPrefix ?? 'asvg'
  const extUri = spec.extUri ?? SVG_BLIP_GUID
  const svgAttr = spec.svg ? (spec.svg.link ? `r:link="svg${spec.id}"` : `r:embed="svg${spec.id}"`) : undefined
  const ext = spec.svg ? `<a:extLst><a:ext uri="${extUri}"><${prefix}:svgBlip xmlns:${prefix}="${ASVG_NS}" ${svgAttr}/></a:ext></a:extLst>` : ''
  return `<p:pic><p:nvPicPr><p:cNvPr id="${spec.id}" name="Picture ${spec.id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip${spec.png ? ` r:embed="png${spec.id}"` : ''}>${ext}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${offsetX}" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
}
function pptxRels(specs: PictureSpec[]): string {
  const rows: string[] = []
  for (const spec of specs) {
    if (spec.png) rows.push(`<Relationship Id="png${spec.id}" Type="${R_NS}/image" Target="${xmlEscape(spec.png.target ?? `../media/${spec.png.name}`)}"${spec.png.external ? ' TargetMode="External"' : ''}/>`)
    if (spec.svg) rows.push(`<Relationship Id="svg${spec.id}" Type="${R_NS}/image" Target="${xmlEscape(spec.svg.target ?? `../media/${spec.svg.name}`)}"${spec.svg.external ? ' TargetMode="External"' : ''}/>`)
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rows.join('')}</Relationships>`
}
export async function buildSvgPptx(pics: PictureSpec[]): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', PPTX_CT)
  zip.file('_rels/.rels', PPTX_ROOT_RELS)
  zip.file('ppt/presentation.xml', PPTX_PRESENTATION)
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/slide" Target="slides/slide1.xml"/></Relationships>`)
  zip.file('ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${pics.map((pic, i) => pptxPic(pic, i)).join('')}</p:spTree></p:cSld></p:sld>`)
  zip.file('ppt/slides/_rels/slide1.xml.rels', pptxRels(pics))
  await writeMedia(zip, 'ppt/media/', pics)
  return zip.generateAsync({ type: 'uint8array' })
}
export async function distinctPptx(): Promise<Uint8Array> {
  return buildSvgPptx(ICONS.map((icon, i) => ({ id: i + 1, png: { name: `icon${i + 1}.png`, data: icon.png }, svg: { name: `icon${i + 1}.svg`, data: icon.svg } })))
}
export async function rasterOnlyPptx(): Promise<Uint8Array> {
  return buildSvgPptx([{ id: 1, png: { name: 'raster.png', data: PNG.purpleSolid } }])
}
export async function fallbackPptx(): Promise<Uint8Array> {
  return buildSvgPptx([
    { id: 1, png: { name: 'missing.png', data: PNG.greenSolid }, svg: { name: 'missing.svg', data: svgBytes(svgDoc('')), target: '../media/absent.svg' } },
    { id: 2, png: { name: 'invalid.png', data: PNG.blueSolid }, svg: { name: 'invalid.svg', data: svgBytes('garbage') } },
    { id: 3, png: { name: 'external.png', data: PNG.orangeSolid }, svg: { name: 'external.svg', data: svgBytes(svgDoc('')), external: true, target: `${CANARY_ORIGIN}/x.svg` } },
    { id: 4, png: { name: 'raster.png', data: PNG.purpleSolid } },
  ])
}
export async function maliciousPptx(): Promise<Uint8Array> {
  return buildSvgPptx(MALICIOUS_MATRIX.map((row, i) => ({ id: i + 1, png: { name: `m${i + 1}.png`, data: PNG.greenSolid }, svg: { name: `m${i + 1}.svg`, data: row.svg } })))
}

// ---- XLSX -----------------------------------------------------------------

const XLSX_CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
const XLSX_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
const XLSX_WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`
const XLSX_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`

function xlsxPic(spec: PictureSpec, index: number): string {
  const width = emu(spec.widthPx ?? 48)
  const height = emu(spec.heightPx ?? 48)
  const prefix = spec.svgPrefix ?? 'asvg'
  const extUri = spec.extUri ?? SVG_BLIP_GUID
  const svgAttr = spec.svg ? (spec.svg.link ? `r:link="svg${spec.id}"` : `r:embed="svg${spec.id}"`) : undefined
  const ext = spec.svg ? `<a:extLst><a:ext uri="${extUri}"><${prefix}:svgBlip xmlns:${prefix}="${ASVG_NS}" ${svgAttr}/></a:ext></a:extLst>` : ''
  const marker = `<xdr:col>${index}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${index}</xdr:row><xdr:rowOff>0</xdr:rowOff>`
  return `<xdr:oneCellAnchor><xdr:from>${marker}</xdr:from><xdr:ext cx="${width}" cy="${height}"/>` +
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${spec.id}" name="Picture ${spec.id}"/></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip${spec.png ? ` r:embed="png${spec.id}"` : ''}>${ext}</a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`
}
function xlsxRels(specs: PictureSpec[]): string {
  const rows: string[] = []
  for (const spec of specs) {
    if (spec.png) rows.push(`<Relationship Id="png${spec.id}" Type="${R_NS}/image" Target="${xmlEscape(spec.png.target ?? `../media/${spec.png.name}`)}"${spec.png.external ? ' TargetMode="External"' : ''}/>`)
    if (spec.svg) rows.push(`<Relationship Id="svg${spec.id}" Type="${R_NS}/image" Target="${xmlEscape(spec.svg.target ?? `../media/${spec.svg.name}`)}"${spec.svg.external ? ' TargetMode="External"' : ''}/>`)
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rows.join('')}</Relationships>`
}
export async function buildSvgXlsx(pics: PictureSpec[]): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', XLSX_CT)
  zip.file('_rels/.rels', XLSX_ROOT_RELS)
  zip.file('xl/_rels/workbook.xml.rels', XLSX_WORKBOOK_RELS)
  zip.file('xl/styles.xml', XLSX_STYLES)
  zip.file('xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R_NS}"><sheets><sheet name="Draw" sheetId="1" r:id="rId1"/></sheets></workbook>`)
  zip.file('xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rDraw"/></worksheet>`)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDraw" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
  zip.file('xl/drawings/drawing1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}" xmlns:asvg="${ASVG_NS}">${pics.map((pic, i) => xlsxPic(pic, i)).join('')}</xdr:wsDr>`)
  zip.file('xl/drawings/_rels/drawing1.xml.rels', xlsxRels(pics))
  await writeMedia(zip, 'xl/media/', pics)
  return zip.generateAsync({ type: 'uint8array' })
}
export async function distinctXlsx(): Promise<Uint8Array> {
  return buildSvgXlsx(ICONS.map((icon, i) => ({ id: i + 1, png: { name: `icon${i + 1}.png`, data: icon.png }, svg: { name: `icon${i + 1}.svg`, data: icon.svg } })))
}
export async function rasterOnlyXlsx(): Promise<Uint8Array> {
  return buildSvgXlsx([{ id: 1, png: { name: 'raster.png', data: PNG.purpleSolid } }])
}
export async function fallbackXlsx(): Promise<Uint8Array> {
  return buildSvgXlsx([
    { id: 1, png: { name: 'missing.png', data: PNG.greenSolid }, svg: { name: 'missing.svg', data: svgBytes(svgDoc('')), target: '../media/absent.svg' } },
    { id: 2, png: { name: 'invalid.png', data: PNG.blueSolid }, svg: { name: 'invalid.svg', data: svgBytes('garbage') } },
    { id: 3, png: { name: 'external.png', data: PNG.orangeSolid }, svg: { name: 'external.svg', data: svgBytes(svgDoc('')), external: true, target: `${CANARY_ORIGIN}/x.svg` } },
    { id: 4, png: { name: 'raster.png', data: PNG.purpleSolid } },
  ])
}
export async function maliciousXlsx(): Promise<Uint8Array> {
  return buildSvgXlsx(MALICIOUS_MATRIX.map((row, i) => ({ id: i + 1, png: { name: `m${i + 1}.png`, data: PNG.greenSolid }, svg: { name: `m${i + 1}.svg`, data: row.svg } })))
}
