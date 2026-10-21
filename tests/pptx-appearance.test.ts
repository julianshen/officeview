import { describe, expect, it, test } from 'vitest';
import JSZip from 'jszip';
import { createCanvas, loadImage } from 'canvas';
import { OfficePackage } from '../src/core/zip';
import { parsePptx } from '../src/pptx/parse';
import { renderSlide } from '../src/pptx/render';
import { parseXmlOrdered } from '../src/core/xml';
import { resolveDrawingStyle, parseThemeContext } from '../src/drawing/style';
import { getPaintables } from '../src/render/paint';
import { buildTextIndex } from '../src/core/search';
import { handleWorkerRequest } from '../src/worker/office-worker';
import { buildPptxDeck, type PptxDeckSpec } from '../src/testdata/ooxml-builders';
import { enrichSlideAppearance } from '../src/index';
import type { PptxDocument } from '../src/pptx/types';
const emu = (n: number) => n * 9525;
const xfrm = (x = 0, y = 0, w = 100, h = 80) => `<a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>`;
const solid = (hex: string, alpha?: number) => `<a:solidFill><a:srgbClr val="${hex}">${alpha === undefined ? '' : `<a:alpha val="${alpha}"/>`}</a:srgbClr></a:solidFill>`;
const shape = (pr: string, x = 0, y = 0, w = 100, h = 80, tag = 'sp', style = '') => `<p:${tag}><p:spPr>${xfrm(x, y, w, h)}<a:prstGeom prst="${tag === 'cxnSp' ? 'line' : 'rect'}"/>${pr}</p:spPr>${style}</p:${tag}>`;
const pic = (effects: string, x: number, crop = '') => `<p:pic><p:blipFill><a:blip r:embed="im">${effects}</a:blip>${crop}</p:blipFill><p:spPr>${xfrm(x, 0, 40, 40)}</p:spPr></p:pic>`;
const cell = (pr = '', attributes = '', text = '') => `<a:tc ${attributes}><a:txBody><a:bodyPr/><a:p>${text ? `<a:r><a:rPr sz="1000"/><a:t>${text}</a:t></a:r>` : ''}</a:p></a:txBody><a:tcPr ${pr.startsWith('mar') ? pr : ''}>${pr.startsWith('mar') ? '' : pr}</a:tcPr></a:tc>`;
const table = (rows: string[], cols = 2, attributes = '') => `<p:graphicFrame><p:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(rows.length * 40)}"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblPr ${attributes}><a:tableStyleId>S</a:tableStyleId></a:tblPr><a:tblGrid>${'<a:gridCol w="476250"/>'.repeat(cols)}</a:tblGrid>${rows.map(r => `<a:tr h="381000">${r}</a:tr>`).join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
const styleParts = { 'ppt/tableStyles.xml': `<a:tblStyleLst><a:tblStyle styleId="S"><a:wholeTbl><a:tcStyle><a:fill>${solid('0000FF')}</a:fill><a:tcBdr><a:insideV><a:ln w="38100">${solid('0000FF')}</a:ln></a:insideV><a:top><a:ln w="38100">${solid('0000FF')}</a:ln></a:top></a:tcBdr></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>` };
async function fixture(body: string, parts: Record<string, string | Uint8Array> = {}) {
    const zip = new JSZip();
    zip.file('ppt/presentation.xml', '<p:presentation><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>');
    zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="s" Type="slide" Target="slides/slide1.xml"/></Relationships>');
    zip.file('ppt/slides/slide1.xml', `<p:sld><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`);
    for (const [path, value] of Object.entries(parts))
        zip.file(path, value);
    return parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })));
}
async function render(doc: Awaited<ReturnType<typeof fixture>>, alpha = 1) {
    const ctx = createCanvas(300, 200).getContext('2d');
    ctx.globalAlpha = alpha;
    const images = await Promise.all(doc.images.map(i => loadImage(Buffer.from(i.data))));
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, undefined, images as unknown as CanvasImageSource[]);
    return { ctx, pixel: (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data) };
}
const theme = parseThemeContext('<a:theme><a:themeElements><a:fmtScheme><a:lnStyleLst><a:ln w="28575" cap="rnd">' + solid('FF0000') + '<a:tailEnd type="triangle"/></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>');
const lnRef = parseXmlOrdered('<p:style><a:lnRef idx="1"/></p:style>');
describe('PPTX native appearance corrections', () => {
    test.each(['native', 'fallback'] as const)('table appearance diagnostics retain %s frame provenance', async representation => {
        const frame = table([cell(solid('GGGGGG') + '<a:lnL w="bad"/>')], 1)
            .replace('<p:graphicFrame>', '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Audit table"/></p:nvGraphicFramePr>')
            .replace('<a:tcPr >', '<a:tcPr marL="bad">');
        const body = representation === 'native' ? frame : `<mc:AlternateContent><mc:Choice Requires="a14">${shape(solid('FF0000'))}</mc:Choice><mc:Fallback>${frame}</mc:Fallback></mc:AlternateContent>`;
        const doc = await fixture(body);
        const parsed = doc.slides[0].shapes[0];
        expect(parsed.source).toMatchObject({ partPath: 'ppt/slides/slide1.xml', element: 'graphicFrame', id: '8', name: 'Audit table', representation });
        for (const kind of ['invalid-color', 'invalid-line', 'invalid-paint']) {
            const shapeIssue = parsed.diagnostics?.find(issue => issue.kind === kind);
            const slideIssue = doc.slides[0].diagnostics?.find(issue => issue.kind === kind);
            expect(shapeIssue).toBeDefined();
            expect(slideIssue).toBeDefined();
            expect(shapeIssue?.source).toEqual(parsed.source);
            expect(slideIssue?.source).toEqual(parsed.source);
        }
    });
    test('empty unstyled line has no invented paint', () => {
        expect(resolveDrawingStyle(parseXmlOrdered('<p:spPr><a:noFill/><a:ln/></p:spPr>')).line?.fill).toBeUndefined();
    });
    test('empty themed line retains actual referenced fields', () => {
        expect(resolveDrawingStyle(parseXmlOrdered('<p:spPr><a:ln/></p:spPr>'), lnRef, theme).line).toMatchObject({ width: 3, cap: 'round', fill: { color: { r: 255 } }, tailEnd: { type: 'triangle' } });
    });
    test('partial width override inherits default paint without adding unstyled paint', () => {
        const pr = parseXmlOrdered('<p:spPr><a:ln w="47625"/></p:spPr>');
        expect(resolveDrawingStyle(pr, lnRef, theme).line).toMatchObject({ width: 5, cap: 'round', fill: { color: { r: 255 } } });
        expect(resolveDrawingStyle(pr).line?.fill).toBeUndefined();
        expect(resolveDrawingStyle(pr, undefined, undefined, { line: { fill: { kind: 'none' }, cap: 'square' } }).line).toMatchObject({ width: 5, cap: 'square', fill: { kind: 'none' } });
    });
    test('removes empty textbox outlines while keeping an adjacent explicit connector', async () => {
        const result = await render(await fixture(shape('<a:noFill/><a:ln/>') + shape(`<a:ln w="38100">${solid('FF0000')}<a:tailEnd type="triangle"/></a:ln>`, 120, 40, 80, 0, 'cxnSp')));
        expect(result.pixel(0, 40)).toEqual([255, 255, 255, 255]);
        expect(result.pixel(160, 40)[0]).toBe(255);
        expect(result.pixel(160, 40)[1]).toBe(0);
    });
    test('repeated image part retains per-use crop, ordered multiplicative opacity, group placement and source order', async () => {
        const c = createCanvas(40, 40), ctx = c.getContext('2d');
        ctx.fillStyle = 'red';
        ctx.fillRect(0, 0, 20, 40);
        ctx.fillStyle = '#00ff00';
        ctx.fillRect(20, 0, 20, 40);
        const groupTransform = xfrm(20, 50, 160, 80).replace('</a:xfrm>', '<a:chOff x="0" y="0"/><a:chExt cx="762000" cy="381000"/></a:xfrm>');
        const doc = await fixture(`<p:grpSp><p:grpSpPr>${groupTransform}</p:grpSpPr>${pic('<a:alphaModFix amt="50000"/><a:alphaModFix amt="50000"/>', 0, '<a:srcRect r="50000"/>')}${pic('<a:alphaModFix amt="50000"/>', 40, '<a:srcRect l="50000"/>')}</p:grpSp>`, { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship Id="im" Type="x/image" Target="../media/i.png"/></Relationships>', 'ppt/media/i.png': new Uint8Array(c.toBuffer('image/png')) });
        const uses = doc.slides[0].shapes[0].children!;
        expect(doc.images).toHaveLength(1);
        expect(uses.map(s => s.imageIndex)).toEqual([0, 0]);
        expect(uses.map(s => s.image?.opacity)).toEqual([.25, .5]);
        expect(doc.images[0].opacity).toBeUndefined();
        const result = await render(doc);
        expect(result.pixel(40, 70)).toEqual([255, 191, 191, 255]);
        expect(result.pixel(120, 70)).toEqual([127, 255, 127, 255]);
        expect(result.pixel(10, 70)).toEqual([255, 255, 255, 255]);
        const caller = await render(doc, .5);
        expect(caller.ctx.globalAlpha).toBe(.5);
        // .125 red over .5 white: premultiplied result is approximately (255,198,198,.5625).
        expect(Math.abs(caller.pixel(40, 70)[3] - 143)).toBeLessThanOrEqual(1);
        expect(Math.abs(caller.pixel(40, 70)[1] - 198)).toBeLessThanOrEqual(1);
    });
    test.each(['0', 'NaN', 'Infinity', '', '-1', '0.5', '0xC350', '5e4', '50000.0'])('preserves zero and diagnoses invalid fixed-alpha amounts %s', async (amount) => {
        const c = createCanvas(2, 2), ctx = c.getContext('2d');
        ctx.fillStyle = 'red';
        ctx.fillRect(0, 0, 2, 2);
        const doc = await fixture(pic(`<a:alphaModFix amt="${amount}"/>`, 0), { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship Id="im" Type="x/image" Target="../media/i.png"/></Relationships>', 'ppt/media/i.png': new Uint8Array(c.toBuffer('image/png')) });
        if (amount === '0') {
            expect(doc.slides[0].shapes[0].image?.opacity).toBe(0);
            expect((await render(doc)).pixel(20, 20)).toEqual([255, 255, 255, 255]);
        }
        else
            expect(doc.slides[0].diagnostics?.some(d => d.kind === 'invalid-paint' && d.feature === 'alphaModFix')).toBe(true);
    });
    test('valid alpha amplification remains diagnosed as unsupported', async () => {
        const c = createCanvas(2, 2);
        const doc = await fixture(pic('<a:alphaModFix amt="200000"/>', 0), { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship Id="im" Type="x/image" Target="../media/i.png"/></Relationships>', 'ppt/media/i.png': new Uint8Array(c.toBuffer('image/png')) });
        expect(doc.slides[0].diagnostics?.some(d => d.kind === 'unsupported-effect' && d.feature === 'alphaModFix')).toBe(true);
    });
    test('omitted fixed-alpha amount uses its schema default', async () => {
        const c = createCanvas(2, 2);
        const doc = await fixture(pic('<a:alphaModFix/>', 0), { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship Id="im" Type="x/image" Target="../media/i.png"/></Relationships>', 'ppt/media/i.png': new Uint8Array(c.toBuffer('image/png')) });
        expect(doc.slides[0].shapes[0].image?.opacity).toBe(1);
        expect(doc.slides[0].diagnostics).toEqual([]);
    });
    test('retains unsupported image effect diagnostics', async () => {
        const c = createCanvas(2, 2);
        const doc = await fixture(pic('<a:grayscl/>', 0), { 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship Id="im" Type="x/image" Target="../media/i.png"/></Relationships>', 'ppt/media/i.png': new Uint8Array(c.toBuffer('image/png')) });
        expect(doc.slides[0].diagnostics?.some(d => d.kind === 'unsupported-effect' && d.feature === 'grayscl')).toBe(true);
    });
    test.each([solid('FFFFFF', 0), '<a:noFill/>'])('explicit transparent/no-fill cell overrides style fill %s', async (fill) => {
        const doc = await fixture(shape(solid('00FF00')) + table([cell(fill) + cell()]), styleParts);
        const result = await render(doc);
        expect(result.pixel(20, 20)).toEqual([0, 255, 0, 255]);
        expect(result.pixel(70, 20)).toEqual([0, 0, 255, 255]);
    });
    test('cell fill and direct borders use shared RGBA color transforms', async () => {
        const doc = await fixture(table([cell(solid('FF0000', 50000) + `<a:lnT w="38100">${solid('00FF00', 50000)}</a:lnT>`) + cell()]), styleParts);
        expect(doc.slides[0].shapes[0].table!.rows[0].cells[0].drawingBorders?.top?.fill).toEqual({ kind: 'solid', color: { r: 0, g: 255, b: 0, a: .5 } });
        const result = await render(doc);
        expect(result.pixel(20, 20)[0]).toBe(255);
        expect(Math.abs(result.pixel(20, 20)[1] - 128)).toBeLessThanOrEqual(1);
        expect(Math.abs(result.pixel(20, 20)[2] - 128)).toBeLessThanOrEqual(1);
        expect(result.pixel(20, 1)[1]).toBeGreaterThan(170);
        expect(result.pixel(20, 1)[2]).toBeLessThan(100);
    });
    test.each(['t', 'ctr', 'b'] as const)('cell margins retain explicit zero and vertical anchor %s', async anchor => {
        const doc = await fixture(table([cell(`marL="0" marR="0" marT="0" marB="0" anchor="${anchor}"`, '', 'x')], 1));
        const parsed = doc.slides[0].shapes[0].table!.rows[0].cells[0];
        expect(parsed.margins).toEqual({ leftEmu: 0, rightEmu: 0, topEmu: 0, bottomEmu: 0 });
        expect(parsed.anchor).toBe(anchor);
        const result = await render(doc);
        let text: [
            string,
            number,
            number
        ][] = [];
        const old = result.ctx.fillText.bind(result.ctx);
        result.ctx.fillText = (t, x, y) => { text.push([t, x, y]); old(t, x, y); };
        // Compare against an independently painted top anchor. Anchor deltas
        // are fixed by the 40px cell and source 10pt normal line advance, not installed fonts.
        const topDoc = await fixture(table([cell('marL="0" marR="0" marT="0" marB="0" anchor="t"', '', 'x')], 1));
        const topCtx = createCanvas(300, 200).getContext('2d');
        let topBaseline = 0;
        topCtx.fillText = (_text, _x, baseline) => { topBaseline = baseline; };
        renderSlide(topDoc.slides[0], topCtx as unknown as CanvasRenderingContext2D);
        renderSlide(doc.slides[0], result.ctx as unknown as CanvasRenderingContext2D);
        expect(text).toHaveLength(1);
        expect(text[0].slice(0, 2)).toEqual(['x', 0]);
        expect(topBaseline).toBeGreaterThan(0);
        expect(topBaseline).toBeLessThan(10 * 96 / 72);
        const shift = (40 - 10 * 96 / 72 * 1.2) * (anchor === 'ctr' ? .5 : anchor === 'b' ? 1 : 0);
        expect(text[0][2] - topBaseline).toBeCloseTo(shift, 10);
    });
    test('direct border precedence and noFill keep merged interiors suppressed', async () => {
        const doc = await fixture(table([cell(solid('FFFFFF') + `<a:lnR w="38100">${solid('FF0000')}</a:lnR>`) + cell(solid('FFFFFF') + '<a:lnL><a:noFill/></a:lnL>'), cell(solid('FFFFFF') + `<a:lnR w="38100">${solid('FF0000')}</a:lnR>`, 'gridSpan="2"') + cell(`<a:lnL w="38100">${solid('00FF00')}</a:lnL>`, 'hMerge="1"')]), styleParts);
        const result = await render(doc);
        expect(result.pixel(50, 20)).toEqual([255, 255, 255, 255]);
        expect(result.pixel(50, 60)).toEqual([255, 255, 255, 255]);
    });
    test('partial direct border width inherits the selected style color', async () => {
        const doc = await fixture(table([cell(solid('FFFFFF') + '<a:lnR w="76200"/>') + cell(solid('FFFFFF'))]), styleParts);
        const result = await render(doc);
        expect(result.pixel(47, 20)).toEqual([0, 0, 255, 255]);
        expect(result.pixel(44, 20)).toEqual([255, 255, 255, 255]);
    });
    test('legacy cell fill and paragraphs remain usable without richer properties', async () => {
        const doc = await fixture(table([cell('', '', 'legacy')], 1), styleParts);
        const legacy = doc.slides[0].shapes[0].table!.rows[0].cells[0];
        legacy.fill = '#00FF00';
        delete legacy.drawingFill;
        expect((await render(doc)).pixel(10, 10)).toEqual([0, 255, 0, 255]);
        expect(legacy.paragraphs[0].runs[0].text).toBe('legacy');
    });
    test('visible direct border replaces the style border', async () => {
        const result = await render(await fixture(table([cell(`<a:lnR w="38100">${solid('FF0000')}</a:lnR>`) + cell()]), styleParts));
        expect(result.pixel(50, 20)).toEqual([255, 0, 0, 255]);
    });
});

/* Slide appearance enrichment (slideshow S0-7). Office-shaped synthetic decks from buildPptxDeck, not Office evidence. */
const pngOf = (rgb: string, size = 4) => {
    const c = createCanvas(size, size), x = c.getContext('2d');
    x.fillStyle = `#${rgb}`;
    x.fillRect(0, 0, size, size);
    return new Uint8Array(c.toBuffer('image/png'));
};
const LOGO = pngOf('00FF00');
const PHOTO = pngOf('0000FF');
// 9144000 x 6858000 EMU = 960 x 720 px; 9525 EMU per px.
const px = (n: number) => String(n * 9525);
function builderDeck(extra: Partial<PptxDeckSpec> = {}): PptxDeckSpec {
    return {
        masters: [{
            bg: { gradient: { angle: 0, stops: [[0, 'FF0000'], [100000, '0000FF']] } },
            shapes: [
                { id: 10, name: 'Master Logo', image: { data: LOGO, name: 'logo.png' }, off: [px(20), px(20)], ext: [px(100), px(100)] },
                { id: 11, name: 'Master Note', fill: 'FFFFFF', off: [px(400), px(600)], ext: [px(300), px(60)], paragraphs: [{ runs: [{ text: 'MasterOnlyText', sz: '2000' }] }] },
                { id: 12, name: 'Master Title', placeholder: { type: 'title' }, off: [px(0), px(0)], ext: [px(960), px(100)] },
            ],
            layouts: [{ shapes: [{ id: 20, name: 'Layout Band', fill: '336699', off: [px(0), px(690)], ext: [px(960), px(30)] }] }],
        }],
        slides: [{ shapes: [{ id: 3, name: 'Slide Box', fill: 'FFFF00', off: [px(400), px(300)], ext: [px(100), px(100)], paragraphs: [{ runs: [{ text: 'SlideText', sz: '1200' }] }] }] }],
        ...extra,
    };
}
const parseDeck = async (spec: PptxDeckSpec) => parsePptx(await OfficePackage.load(await buildPptxDeck(spec)));
async function paintFirst(doc: PptxDocument, options: Parameters<typeof getPaintables>[1] = {}) {
    const pages = await getPaintables(doc, options);
    const page = pages[0];
    const canvas = createCanvas(page.spec.widthPx, page.spec.heightPx), ctx = canvas.getContext('2d');
    page.paint(ctx as unknown as CanvasRenderingContext2D);
    pages.dispose();
    return { pages, ctx, pixel: (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data) };
}
const MASTER_PART = 'ppt/slideMasters/slideMaster1.xml';
/** Append raw XML to a part's spTree (the builder has no useBgFill or scheme-colour shapes). */
async function patchPart(bytes: Uint8Array, part: string, xml: string): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(bytes);
    zip.file(part, (await zip.file(part)!.async('text')).replace('</p:spTree>', `${xml}</p:spTree>`));
    return zip.generateAsync({ type: 'uint8array' });
}
const useBgFillShape = (id: number, x: number, y: number, w: number, h: number) => `<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="${id}" name="Bg Window ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
const near = (actual: number[], expected: number[], tolerance = 24) => actual.slice(0, 3).every((v, i) => Math.abs(v - expected[i]) <= tolerance);
const allDiagnostics = (doc: PptxDocument) => [...(doc.diagnostics ?? []), ...doc.slides.flatMap(s => s.diagnostics ?? [])];

describe('slide appearance enrichment', () => {
    it('getPaintables paints a gradient background and a master logo for a builder deck', async () => {
        const { pixel } = await paintFirst(await parseDeck(builderDeck()));
        // linear gradient at 0 degrees: red on the left edge, blue on the right
        expect(near(pixel(2, 360), [255, 0, 0], 40)).toBe(true);
        expect(near(pixel(957, 360), [0, 0, 255], 40)).toBe(true);
        expect(near(pixel(480, 200), [128, 0, 128], 40)).toBe(true);
        // master logo picture, layout band, slide box on top
        expect(near(pixel(70, 70), [0, 255, 0])).toBe(true);
        expect(near(pixel(480, 705), [0x33, 0x66, 0x99])).toBe(true);
        expect(near(pixel(495, 395), [255, 255, 0])).toBe(true);
    });
    it('enrichment is idempotent across two getPaintables calls', async () => {
        const doc = await parseDeck(builderDeck());
        await paintFirst(doc);
        const underlay = doc.slides[0].underlay;
        const images = doc.images.length;
        expect(underlay?.length).toBe(4); // background, master logo, master note, layout band
        await paintFirst(doc);
        expect(doc.slides[0].underlay).toBe(underlay);
        expect(doc.slides[0].underlay?.length).toBe(4);
        expect(doc.images.length).toBe(images);
        // concurrent first calls share one run
        const fresh = await parseDeck(builderDeck());
        await Promise.all([enrichSlideAppearance(fresh), enrichSlideAppearance(fresh)]);
        expect(fresh.slides[0].underlay?.length).toBe(4);
        expect(fresh.images.length).toBe(images);
    });
    it('underlay pictures get an image index and are decoded', async () => {
        const doc = await parseDeck(builderDeck({ slides: [{ shapes: [{ id: 3, name: 'Slide Logo', image: { data: LOGO, name: 'logo.png' }, off: [px(600), px(100)], ext: [px(50), px(50)] }] }] }));
        expect(doc.images).toHaveLength(1);
        await enrichSlideAppearance(doc);
        const logo = doc.slides[0].underlay!.find(s => s.source?.name === 'Master Logo')!;
        // the same media part shared by the slide and the master deduplicates to one asset
        expect(logo.imageIndex).toBe(0);
        expect(doc.images).toHaveLength(1);
        const decoded: string[] = [];
        await paintFirst(doc, { decodeImage: async (bytes, mime) => { decoded.push(mime ?? ''); return loadImage(Buffer.from(bytes)) as never; } });
        expect(decoded).toHaveLength(1);
    });
    it('the watermark paints over an opaque picture background', async () => {
        const doc = await parseDeck(builderDeck({ masters: [{ bg: { picture: { data: PHOTO } }, layouts: [{}] }] }));
        const plain = await paintFirst(doc);
        expect(near(plain.pixel(200, 200), [0, 0, 255])).toBe(true);
        const marked = await paintFirst(doc, { watermark: { text: 'DRAFT', color: '#FF0000', opacity: 1, fontSizePt: 60 } });
        const data = marked.ctx.getImageData(0, 0, 960, 720).data;
        let redOverPicture = 0, redOverBox = 0;
        for (let y = 0; y < 720; y++) for (let x = 0; x < 960; x++) {
            const i = (y * 960 + x) * 4;
            if (data[i] > 200 && data[i + 1] < 60 && data[i + 2] < 60) {
                if (x >= 402 && x < 498 && y >= 302 && y < 398) redOverBox++;
                else redOverPicture++;
            }
        }
        expect(redOverPicture).toBeGreaterThan(500);
        // the slide's own shapes still paint over the watermark
        expect(redOverBox).toBe(0);
    });
    it('renderSlide without enrichment paints the pre-Stage-0 appearance; after enrichSlideAppearance it paints the underlay', async () => {
        const doc = await parseDeck(builderDeck());
        const draw = () => {
            const ctx = createCanvas(960, 720).getContext('2d');
            renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D);
            return (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data);
        };
        const before = draw();
        expect(before(2, 360)).toEqual([255, 255, 255, 255]);
        expect(before(480, 705)).toEqual([255, 255, 255, 255]);
        expect(doc.slides[0].underlay).toBeUndefined();
        await enrichSlideAppearance(doc);
        const after = draw();
        expect(near(after(2, 360), [255, 0, 0], 40)).toBe(true);
        expect(near(after(480, 705), [0x33, 0x66, 0x99])).toBe(true);
        // the solid approximation is unchanged for callers that read it
        expect(doc.slides[0].background).toBe(undefined);
    });
    it('the worker path (handleWorkerRequest) returns the enriched slide', async () => {
        const source = await buildPptxDeck(builderDeck());
        const env = {
            async createCanvas(w: number, h: number) { return createCanvas(w, h); },
            async toBitmap(canvas: unknown) { return canvas as ImageBitmap; },
            async parse(_format: 'pptx', bytes: Uint8Array) { return parsePptx(await OfficePackage.load(bytes)); },
            async paintables(doc: never, options: never) { return getPaintables(doc, options); },
        };
        const res = await handleWorkerRequest({ id: 7, kind: 'render-slide', format: 'pptx', source: source.slice().buffer as ArrayBuffer, index: 0 }, env as never);
        expect(res.ok).toBe(true);
        if (!res.ok) return;
        const ctx = (res.bitmap as unknown as ReturnType<typeof createCanvas>).getContext('2d');
        expect(near(Array.from(ctx.getImageData(2, 360, 1, 1).data), [255, 0, 0], 40)).toBe(true);
        expect(near(Array.from(ctx.getImageData(70, 70, 1, 1).data), [0, 255, 0])).toBe(true);
    });
    it('a cloned document without a package handle paints as before, with a no-package diagnostic', async () => {
        const clone = structuredClone(await parseDeck(builderDeck()));
        const { pixel } = await paintFirst(clone);
        expect(pixel(2, 360)).toEqual([255, 255, 255, 255]);
        expect(clone.slides[0].underlay).toBeUndefined();
        await paintFirst(clone);
        expect(allDiagnostics(clone).filter(d => d.kind === 'no-package')).toHaveLength(1);
    });
    it('an invalid slide size skips enrichment with a slide-size diagnostic', async () => {
        const doc = await parseDeck(builderDeck({ slideSize: [9144000, 100] }));
        await enrichSlideAppearance(doc);
        expect(doc.slides[0].underlay).toBeUndefined();
        expect(allDiagnostics(doc).filter(d => d.kind === 'slide-size')).toHaveLength(1);
        await enrichSlideAppearance(doc);
        expect(allDiagnostics(doc).filter(d => d.kind === 'slide-size')).toHaveLength(1);
    });
    it('an appearance chunk failure still returns pages', async () => {
        const doc = await parseDeck(builderDeck());
        const { pages, pixel } = await paintFirst(doc, { __appearanceLoader: () => Promise.reject(new Error('chunk load failed')) });
        expect(pages).toHaveLength(1);
        expect(pixel(2, 360)).toEqual([255, 255, 255, 255]);
        expect(near(pixel(495, 395), [255, 255, 0])).toBe(true);
        expect(allDiagnostics(doc).filter(d => d.kind === 'appearance-unavailable')).toHaveLength(1);
    });
    it('the effective slide colour map reaches a bgRef background (phClr)', async () => {
        // bgRef with phClr from a scheme colour: bg1 maps to dk1 (black) through the slide clrMapOvr
        const spec = builderDeck({
            masters: [{ bg: { bgRef: { idx: 1001, color: 'bg1' } }, layouts: [{}] }],
            slides: [{ shapes: [], extraXml: { afterCSld: '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>' } }],
        });
        const doc = await parseDeck(spec);
        await enrichSlideAppearance(doc);
        const bg = doc.slides[0].underlay?.[0];
        expect(bg?.drawingStyle?.fill?.kind).toBe('solid');
        const color = (bg!.drawingStyle!.fill as { color: { r: number; g: number; b: number } }).color;
        expect(color.r + color.g + color.b).toBeLessThan(120);
    });
    it('the effective slide colour map reaches master artwork colours (two slides, two maps)', async () => {
        // a master shape filled with schemeClr tx1: dk1 (black) under the master map, lt1 (white) under the override
        const box = `<p:sp><p:nvSpPr><p:cNvPr id="41" name="Scheme Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${px(300)}" y="${px(300)}"/><a:ext cx="${px(100)}" cy="${px(100)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></p:spPr></p:sp>`;
        const override = '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>';
        const bytes = await patchPart(await buildPptxDeck(builderDeck({
            masters: [{ bg: { solid: '808080' }, layouts: [{}] }],
            slides: [{ shapes: [] }, { shapes: [], extraXml: { afterCSld: override } }],
        })), MASTER_PART, box);
        const doc = await parsePptx(await OfficePackage.load(bytes));
        const pages = await getPaintables(doc);
        const pixelOf = (index: number) => {
            const ctx = createCanvas(960, 720).getContext('2d');
            pages[index].paint(ctx as unknown as CanvasRenderingContext2D);
            return Array.from(ctx.getImageData(350, 350, 1, 1).data);
        };
        expect(near(pixelOf(0), [0, 0, 0], 8)).toBe(true);
        expect(near(pixelOf(1), [255, 255, 255], 8)).toBe(true);
        pages.dispose();
    });
    it('a useBgFill shape in a shared master shows each slide\'s own background', async () => {
        const bytes = await patchPart(await buildPptxDeck(builderDeck({
            masters: [{ shapes: [], layouts: [{}] }],
            slides: [{ bg: { solid: 'FF0000' }, shapes: [] }, { bg: { solid: '0000FF' }, shapes: [] }],
        })), MASTER_PART, `<p:sp><p:nvSpPr><p:cNvPr id="40" name="Under"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:spPr></p:sp>` + useBgFillShape(42, 50, 50, 300, 300));
        const doc = await parsePptx(await OfficePackage.load(bytes));
        const pages = await getPaintables(doc);
        const at = (index: number, x: number, y: number) => {
            const ctx = createCanvas(960, 720).getContext('2d');
            pages[index].paint(ctx as unknown as CanvasRenderingContext2D);
            return Array.from(ctx.getImageData(x, y, 1, 1).data);
        };
        // inside the useBgFill shape: the slide's background, through the master's green rectangle
        expect(near(at(0, 100, 100), [255, 0, 0], 8)).toBe(true);
        expect(near(at(1, 100, 100), [0, 0, 255], 8)).toBe(true);
        // outside it: the master's green rectangle
        expect(near(at(0, 600, 600), [0, 255, 0], 8)).toBe(true);
        pages.dispose();
    });
    it('a useBgFill shape takes a gradient background in slide coordinates', async () => {
        // a full-slide white master rectangle (lo-tdf93868's shape) hides the gradient except through the slide's useBgFill shape
        const bytes = await patchPart(await patchPart(await buildPptxDeck(builderDeck({
            masters: [{ bg: { gradient: { angle: 0, stops: [[0, 'FF0000'], [100000, '0000FF']] } }, shapes: [{ id: 10, name: 'White', fill: 'FFFFFF', off: ['0', '0'], ext: ['9144000', '6858000'] }], layouts: [{}] }],
            slides: [{ shapes: [] }],
        })), MASTER_PART, ''), 'ppt/slides/slide1.xml', useBgFillShape(43, 760, 100, 160, 200));
        const { pixel } = await paintFirst(await parsePptx(await OfficePackage.load(bytes)));
        expect(pixel(100, 360)).toEqual([255, 255, 255, 255]);
        // slide-coordinate gradient: at x=770 the slide gradient is ~80% of the way to blue, not the shape's red start
        const left = pixel(770, 200), right = pixel(910, 200);
        expect(near(left, [Math.round(255 * (1 - 770 / 960)), 0, Math.round(255 * 770 / 960)], 24), `left ${left}`).toBe(true);
        expect(near(right, [Math.round(255 * (1 - 910 / 960)), 0, Math.round(255 * 910 / 960)], 24), `right ${right}`).toBe(true);
    });
    it('a useBgFill shape takes a picture background in slide coordinates', async () => {
        const c = createCanvas(20, 10), x = c.getContext('2d');
        x.fillStyle = '#FF0000'; x.fillRect(0, 0, 10, 10);
        x.fillStyle = '#00FF00'; x.fillRect(10, 0, 10, 10);
        const halves = new Uint8Array(c.toBuffer('image/png'));
        const bytes = await patchPart(await buildPptxDeck(builderDeck({
            masters: [{ bg: { picture: { data: halves } }, shapes: [{ id: 10, name: 'White', fill: 'FFFFFF', off: ['0', '0'], ext: ['9144000', '6858000'] }], layouts: [{}] }],
            slides: [{ shapes: [] }],
        })), 'ppt/slides/slide1.xml', useBgFillShape(44, 100, 100, 200, 100) + useBgFillShape(45, 660, 100, 200, 100));
        const { pixel } = await paintFirst(await parsePptx(await OfficePackage.load(bytes)));
        expect(near(pixel(200, 150), [255, 0, 0], 8)).toBe(true);
        expect(near(pixel(760, 150), [0, 255, 0], 8)).toBe(true);
        expect(pixel(480, 400)).toEqual([255, 255, 255, 255]);
    });
    it('a master shared by several slides is parsed once per run (one node-budget charge)', async () => {
        const slide = { shapes: [{ id: 3, name: 'Box', fill: 'FFFF00', off: [px(400), px(300)] as [string, string], ext: [px(100), px(100)] as [string, string] }] };
        const doc = await parseDeck(builderDeck({ slides: [slide, slide, slide] }));
        await enrichSlideAppearance(doc);
        const logos = doc.slides.map(s => s.underlay!.find(n => n.source?.name === 'Master Logo'));
        expect(logos[0]).toBeDefined();
        expect(logos[1]).toBe(logos[0]);
        expect(logos[2]).toBe(logos[0]);
        // each slide still gets its own underlay list and background node
        expect(doc.slides[1].underlay).not.toBe(doc.slides[0].underlay);
        expect(doc.images.filter(i => i.partPath === logos[0]!.image!.partPath)).toHaveLength(1);
    });
    it('text search and coverage ignore underlay shapes', async () => {
        const doc = await parseDeck(builderDeck());
        const pages = await getPaintables(doc);
        expect(doc.slides[0].underlay?.some(s => s.textBody?.paragraphs.some(p => p.runs.some(r => r.text === 'MasterOnlyText')))).toBe(true);
        const index = await buildTextIndex(pages);
        const text = index.pages[0].lines.map(line => line.text).join(' ');
        expect(text).toContain('SlideText');
        expect(text).not.toContain('MasterOnlyText');
        expect(doc.drawingCoverage?.every(entry => entry.partPath === 'ppt/slides/slide1.xml')).toBe(true);
        pages.dispose();
    });
});
