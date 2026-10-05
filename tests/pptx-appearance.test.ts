import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { createCanvas, loadImage } from 'canvas';
import { OfficePackage } from '../src/core/zip';
import { parsePptx } from '../src/pptx/parse';
import { renderSlide } from '../src/pptx/render';
import { parseXmlOrdered } from '../src/core/xml';
import { resolveDrawingStyle, parseThemeContext } from '../src/drawing/style';
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
