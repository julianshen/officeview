import { describe, it, expect } from 'vitest';
import { buildPptxDeck } from '../src/testdata/ooxml-builders';
import { createDiagnosticSink } from '../src/slideshow/diagnostics';
import type { PptxDocument } from '../src/pptx/types';
const doc = (n: number): PptxDocument => ({ slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: Array.from({ length: n }, (_, index) => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) });
const props = (body: string, attrs = '') => `<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr ${attrs}>${body}</p:showPr></p:presentationPr>`;
async function read(xml?: string, hidden = false) {
    const { readShowIndex, openShowPackage } = await import('../src/slideshow/read/deck');
    const handle = await openShowPackage(await buildPptxDeck({ slides: [{ shapes: [] }, { shapes: [], hidden }, { shapes: [] }], presProps: xml }));
    const sink = createDiagnosticSink();
    return { index: await readShowIndex(doc(3), handle, sink), sink, handle };
}
describe('show index', () => {
    it('showPr defaults when attributes are absent; kiosk restart default 300000', async () => {
        expect((await read(props('<p:present/>'))).index.settings).toEqual({ loop: false, showAnimation: true, showNarration: false, useTimings: true, type: 'present', kioskRestartMs: 300000 });
        expect((await read(props('<p:kiosk/>'))).index.settings).toMatchObject({ type: 'kiosk', kioskRestartMs: 300000 });
    });
    it('sldRg st..end inclusive; sldRg outside the deck falls back to the whole deck', async () => {
        expect((await read(props('<p:sldRg st="2" end="3"/>'))).index.settings.range).toEqual({ st: 1, end: 2 });
        for (const range of ['0:2', '2:4', '3:2']) {
            const [st, end] = range.split(':');
            const r = await read(props(`<p:sldRg st="${st}" end="${end}"/>`));
            expect(r.index.settings.range).toBeUndefined();
            expect(r.sink.list().some(d => d.code === 'slide-range')).toBe(true);
        }
    });
    it('custShow and browse are diagnosed and play the full deck as present', async () => {
        const r = await read(props('<p:browse/><p:custShow id="1"/>'));
        expect(r.index.settings.type).toBe('present');
        expect(r.index.settings.range).toBeUndefined();
        expect(r.sink.list().map(d => d.code)).toEqual(expect.arrayContaining(['custom-show', 'browse']));
    });
    it('hidden slides keep their index', async () => { expect((await read(undefined, true)).index.slides.map(s => [s.index, s.hidden])).toEqual([[0, false], [1, true], [2, false]]); });
    it('without a handle every slide is static and no-package is diagnosed once', async () => { const { readShowIndex } = await import('../src/slideshow/read/deck'); const sink = createDiagnosticSink(); const index = await readShowIndex(doc(3), undefined, sink); expect(index.slides).toHaveLength(3); expect(sink.list().map(d => d.code)).toEqual(['no-package']); });
});
it('a 1 MiB presProps is read; 1 MiB + 1 is refused with defaults applied', async () => {
    for (const extra of [0, 1]) {
        let x = props('<p:present/>', 'loop="1"');
        x = x.replace('</p:presentationPr>', '<!--' + 'x'.repeat(1024 * 1024 + extra - x.length - 7) + '--></p:presentationPr>');
        const r = await read(x);
        expect(r.index.settings.loop).toBe(extra === 0);
        expect(r.sink.list().some(d => d.code === 'show-properties')).toBe(extra === 1);
    }
});
it('browse ignores a stored range while diagnosing its present-mode fallback', async () => { const r = await read(props('<p:browse/><p:sldRg st="2" end="2"/>')); expect(r.index.settings.range).toBeUndefined(); });
it('presProps is resolved from the package presentation relationship rather than a fixed filename',async()=>{
 const {default:JSZip}=await import('jszip');const zip=await JSZip.loadAsync(await buildPptxDeck({slides:[{shapes:[]}],presProps:props('<p:present/>','loop="1"')}));zip.file('ppt/renamed.xml',await zip.file('ppt/presentation.xml')!.async('string'));zip.file('ppt/_rels/renamed.xml.rels',await zip.file('ppt/_rels/presentation.xml.rels')!.async('string'));zip.remove('ppt/presentation.xml');zip.remove('ppt/_rels/presentation.xml.rels');zip.file('_rels/.rels',(await zip.file('_rels/.rels')!.async('string')).replace('ppt/presentation.xml','ppt/renamed.xml'))
 const {openShowPackage,readShowIndex}=await import('../src/slideshow/read/deck');const h=await openShowPackage(await zip.generateAsync({type:'uint8array'}));expect((await readShowIndex(doc(1),h,createDiagnosticSink())).settings.loop).toBe(true)
})
