import { existsSync } from 'node:fs';
import { it, expect, vi } from 'vitest';
import { buildPptxDeck } from '../src/testdata/ooxml-builders';
import { createDiagnosticSink } from '../src/slideshow/diagnostics';
import { openShowPackage, readShowIndex } from '../src/slideshow/read/deck';
import type { Lease, SlideShowWorkspace, ShowSlide } from '../src/slideshow/types';
import type { PptxDocument } from '../src/pptx/types';
export function workspace(limit = 512 * 1024 * 1024): SlideShowWorkspace {
    let used = 0;
    const owner: SlideShowWorkspace = { limit, used: () => used, reserveTerminal: () => undefined, lease: (bytes, tag) => {
            if (used + bytes > limit)
                return undefined;
            used += bytes;
            let live = true, n = bytes;
            const l: Lease = { get bytes() { return n; }, tag, grow(b) { if (!live || used + b > limit)
                    return false; used += b; n += b; return true; }, shrink(b) { used -= b; n -= b; }, transfer: () => l, release() { if (live) {
                    used -= n;
                    n = 0;
                    live = false;
                } } };
            return l;
        } };
    return owner;
}
async function setup(estimate?: (s: ShowSlide) => number, w = workspace()) {
    const handle = await openShowPackage(await buildPptxDeck({ slides: [{ shapes: [], notes: [['hello']] }, { shapes: [] }] }));
    const doc: PptxDocument = { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: [0, 1].map(index => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) };
    const sink = createDiagnosticSink(), index = await readShowIndex(doc, handle, sink);
    const { createSlideMetadataStore } = await import('../src/slideshow/read/store');
    const timing = vi.fn((_xml: unknown) => ({ timing: undefined, builds: [] }));
    const store = createSlideMetadataStore({ doc, handle, index, workspace: w, sink, timing, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }), estimate });
    return { store, handle, w, timing, sink };
}
it('acquire reads each slide part once and hands the same XML to the timing reader', async () => {
    const { store, handle, timing, w } = await setup();
    const read = vi.spyOn(handle.pkg, 'entry');
    const [a, b] = await Promise.all([store.acquire(0), store.acquire(0)]);
    expect(read.mock.calls.filter(([p]) => p === 'ppt/slides/slide1.xml')).toHaveLength(1);
    expect(timing).toHaveBeenCalledTimes(1);
    expect(timing.mock.calls[0][0]).toHaveProperty('cSld');
    expect(a.slide).toBe(b.slide);
    expect(a.slide.notes[0].text).toBe('hello');
    a.release();
    b.release();
    store.close();
    expect(w.used()).toBe(0);
});
it('evict with no handles releases the charge; a revisited slide is re-read', async () => {
    const { store, handle, w } = await setup();
    const read = vi.spyOn(handle.pkg, 'entry');
    const a = await store.acquire(0);
    a.release();
    store.evict(0);
    expect(w.used()).toBe(0);
    (await store.acquire(0)).release();
    expect(read.mock.calls.filter(([p]) => p === 'ppt/slides/slide1.xml')).toHaveLength(2);
    store.close();
});
it('evicting a slide whose handle an outgoing transition still holds keeps the charge until release; re-acquire reads a fresh entry', async () => {
    const { store, w } = await setup(() => 1000);
    const a = await store.acquire(0);
    store.evict(0);
    expect(w.used()).toBe(1000);
    const b = await store.acquire(0);
    expect(b.slide).not.toBe(a.slide);
    expect(w.used()).toBe(2000);
    a.release();
    a.release();
    expect(w.used()).toBe(1000);
    b.release();
    store.close();
    expect(w.used()).toBe(0);
});
it('a preview handle held across eviction stays valid and is charged once', async () => {
    const { store, w } = await setup(() => 1000);
    const a = await store.acquire(0), b = await store.acquire(0);
    store.evict(0);
    a.release();
    expect(b.slide.notes[0].text).toBe('hello');
    expect(w.used()).toBe(1000);
    store.close();
    expect(w.used()).toBe(1000);
    b.release();
    expect(w.used()).toBe(0);
});
it('metadata at 64 MiB is accepted; one byte more evicts a neighbour, then makes the slide static', async () => {
    const cap = 64 * 1024 * 1024;
    const { store, w } = await setup(s => s.index === 0 ? cap : cap + 1);
    const a = await store.acquire(0);
    expect(w.used()).toBe(cap);
    a.release();
    const b = await store.acquire(1);
    expect(store.resident()).not.toContain(0);
    expect(b.slide.timing).toEqual({ refused: 'metadata-limit' });
    expect(w.used()).toBeLessThan(cap);
    b.release();
    store.close();
});
it('attach that would cross the 64 MiB aggregate or the workspace is refused before allocation', async () => {
    const { store, w } = await setup(() => 1024);
    const a = await store.acquire(0);
    expect(a.attach(64 * 1024 * 1024 - 1024, 'compiled')).toBe(true);
    expect(a.attach(1, 'overflow')).toBe(false);
    expect(w.used()).toBe(64 * 1024 * 1024);
    a.release();
    store.close();
    expect(w.used()).toBe(0);
    const q = await setup(() => 1024, workspace(90 * 1024 * 1024));
    const b = await q.store.acquire(0);
    const pressure = q.w.lease(q.w.limit - 1024, 'pressure')!;
    expect(b.attach(1, 'overflow')).toBe(false);
    pressure.release();
    b.release();
    q.store.close();
});
it('close cancels an in-flight acquire without charging', async () => {
    const { store, w } = await setup();
    const p = store.acquire(0);
    store.close();
    await expect(p).rejects.toMatchObject({ code: 'cancelled' });
    expect(w.used()).toBe(0);
});
async function fixture(spec: Parameters<typeof buildPptxDeck>[0], edit?: (zip: import('jszip')) => Promise<void>) {
    const { default: JSZip } = await import('jszip');
    let data = await buildPptxDeck(spec);
    if (edit) {
        const zip = await JSZip.loadAsync(data);
        await edit(zip);
        data = await zip.generateAsync({ type: 'uint8array' });
    }
    const handle = await openShowPackage(data), sink = createDiagnosticSink(), w = workspace();
    const doc: PptxDocument = { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: spec.slides.map((_, index) => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) };
    const index = await readShowIndex(doc, handle, sink);
    const { createSlideMetadataStore } = await import('../src/slideshow/read/store');
    const store = createSlideMetadataStore({ doc, handle, index, sink, workspace: w, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }) });
    return { store, handle, sink, w, index, doc };
}
it('notes keep level and bullet; 65536 chars of notes are kept and 65537 truncate with a diagnostic; a 1 MiB notes part is read and 1 MiB + 1 is refused', async () => {
    for (const n of [65536, 65537]) {
        const q = await fixture({ slides: [{ shapes: [], notes: [['x'.repeat(n)]] }] }, async (zip) => { const p = 'ppt/notesSlides/notesSlide1.xml'; zip.file(p, (await zip.file(p)!.async('string')).replace('<a:p>', '<a:p><a:pPr lvl="2"><a:buChar char="•"/></a:pPr>')); });
        const h = await q.store.acquire(0);
        expect(h.slide.notes[0]).toMatchObject({ level: 2, bullet: '•', text: 'x'.repeat(65536) });
        expect(q.sink.list().some(d => d.code === 'notes-truncated')).toBe(n > 65536);
        if(n > 65536) expect(h.slide.timing).toEqual({refused:'notes-limit'});
        h.release();
        q.store.close();
    }
    for (const extra of [0, 1]) {
        const q = await fixture({ slides: [{ shapes: [], notes: [['hello']] }] }, async (zip) => { const p = 'ppt/notesSlides/notesSlide1.xml', x = await zip.file(p)!.async('string'); zip.file(p, x.replace('</p:notes>', '<!--' + 'x'.repeat(1024 * 1024 + extra - x.length - 7) + '--></p:notes>')); });
        const h = await q.store.acquire(0);
        expect(h.slide.notes.length).toBe(extra ? 0 : 1);
        expect(q.sink.list().some(d => d.code === 'notes-limit')).toBe(!!extra);
        h.release();
        q.store.close();
    }
});
it('a 4 MiB slide part is read; 4 MiB + 1 drops that slide to static with a diagnostic', async () => {
    for (const extra of [0, 1]) {
        const q = await fixture({ slides: [{ shapes: [] }] }, async (zip) => { const p = 'ppt/slides/slide1.xml', x = await zip.file(p)!.async('string'); zip.file(p, x.replace('</p:sld>', '<!--' + 'x'.repeat(4 * 1024 * 1024 + extra - x.length - 7) + '--></p:sld>')); });
        const h = await q.store.acquire(0);
        expect(h.slide.timing).toEqual(extra ? { refused: 'read-limit' } : undefined);
        h.release();
        q.store.close();
        expect(q.w.used()).toBe(0);
    }
});
it('notes, links, media and an injected timing tree that each fit but together exceed the cap are refused as one aggregate', async () => {
    const q = await fixture({ slides: [{ shapes: [{ id: 2, paragraphs: [{ runs: [{ text: 'link' }] }] }], notes: [['note']], links: [{ shapeId: 2, target: 'https://example.com', external: true }], media: [{ shapeId: 2, kind: 'audio', ext: 'mp3', data: new Uint8Array([1]) }] }] });
    q.store.close();
    const { createSlideMetadataStore } = await import('../src/slideshow/read/store');
    const estimate = vi.fn((s: ShowSlide) => (s.notes.length + s.links.length + s.media.length + (s.timing ? 1 : 0)) * 20 * 1024 * 1024);
    const store = createSlideMetadataStore({ ...q, workspace: q.w, timing: () => ({ timing: { root: { kind: 'par', ctn: { id: 1, durMs: 100, accel: 0, decel: 0, autoRev: false, start: [], end: [] }, children: [] }, build: [] }, builds: [] }), resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }), estimate });
    const h = await store.acquire(0);
    expect(estimate).toHaveReturnedWith(80 * 1024 * 1024);
    expect(h.slide.timing).toEqual({ refused: 'metadata-limit' });
    h.release();
    store.close();
});
it('hlinkClick on shape and run resolves slide jumps, show jumps and http/https/mailto; javascript:, file: and ppaction://macro are refused', async () => {
    const cases = [{ target: 'slide2.xml', action: 'ppaction://hlinksldjump' }, { action: 'ppaction://hlinkshowjump?jump=nextslide' }, ...['http://example.com', 'https://example.com', 'mailto:a@example.com', 'javascript:alert(1)', 'file:///tmp/x'].map(target => ({ target, external: true })), { action: 'ppaction://macro?name=x' }];
    const shapes = cases.map((_, i) => ({ id: i + 2, paragraphs: [{ runs: [{ text: 'link' }] }], off: ['0', '0'] as [
            string,
            string
        ], ext: ['1000', '1000'] as [
            string,
            string
        ] }));
    const q = await fixture({ slides: [{ shapes, links: cases.map((x, i) => ({ shapeId: i + 2, ...x, onRun: i === 2 })) }, { shapes: [] }] });
    const h = await q.store.acquire(0);
    expect(h.slide.links.map(l => l.action.kind)).toEqual(['slide', 'jump', 'url', 'url', 'url', 'refused', 'refused', 'refused']);
    expect(h.slide.links[0].action).toEqual({ kind: 'slide', index: 1 });
    expect(h.slide.links[2].run).toEqual({ p: 0, r: 0 });
    h.release();
    q.store.close();
});
it('media records keep declared size and content type; no bytes are read', async () => {
    const q = await fixture({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }], media: [{ shapeId: 2, kind: 'video', ext: 'mp4', data: new Uint8Array([1, 2, 3]), p14: true }] }] });
    const read = vi.spyOn(q.handle.pkg, 'bytes');
    const streams = q.handle.pkg.paths(p => p.startsWith('ppt/media/')).map(p => vi.spyOn(q.handle.pkg.entry(p) as unknown as {
        internalStream: () => unknown;
    }, 'internalStream'));
    const h = await q.store.acquire(0);
    expect(h.slide.media.length).toBeGreaterThan(0);
    expect(h.slide.media[0]).toMatchObject({ declaredBytes: 3, contentType: 'video/mp4' });
    expect(read).not.toHaveBeenCalled();
    for (const spy of streams)
        expect(spy).not.toHaveBeenCalled();
    h.release();
    q.store.close();
});
it.skipIf(!existsSync('corpus/tdf104788.pptx'))('tdf104788 reads 6 notes bodies and showNarration=1', async () => {
    const { readFileSync } = await import('node:fs');
    const handle = await openShowPackage(new Uint8Array(readFileSync('corpus/tdf104788.pptx')));
    const sink = createDiagnosticSink(), w = workspace();
    const doc: PptxDocument = { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: handle.slideParts.map((_, index) => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) };
    const index = await readShowIndex(doc, handle, sink);
    expect(index.settings.showNarration).toBe(true);
    expect(index.slides).toHaveLength(6);
    const { createSlideMetadataStore } = await import('../src/slideshow/read/store');
    const store = createSlideMetadataStore({ doc, handle, index, sink, workspace: w, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }) });
    for (let i = 0; i < 6; i++) {
        const h = await store.acquire(i);
        expect(h.slide.notes.length).toBeGreaterThan(0);
        h.release();
    }
    store.close();
    expect(w.used()).toBe(0);
});
it('eviction during an acquire prevents cache publication without invalidating its eventual handle', async () => {
    const { store, w } = await setup(() => 1000);
    const p = store.acquire(0);
    store.evict(0);
    const h = await p;
    expect(store.resident()).toEqual([]);
    expect(w.used()).toBe(1000);
    h.release();
    expect(w.used()).toBe(0);
    store.close();
});
it('a released handle cannot regrow and invalid growth never changes accounting', async () => {
    const { store, w } = await setup(() => 1000);
    const h = await store.acquire(0);
    for (const n of [-1, NaN, Infinity, .5])
        expect(h.attach(n, 'bad')).toBe(false);
    expect(w.used()).toBe(1000);
    h.release();
    expect(h.attach(1, 'late')).toBe(false);
    store.close();
});
it('metadata construction is refused before the timing reader when live handles fill the aggregate', async () => {
    const { store, timing } = await setup(() => 64 * 1024 * 1024);
    const a = await store.acquire(0);
    await expect(store.acquire(1)).rejects.toMatchObject({ code: 'resource-limit' });
    expect(timing).toHaveBeenCalledTimes(1);
    a.release();
    store.close();
});
