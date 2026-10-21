import { parseXmlOrdered, type XmlNode } from '../../core/xml';
import { PackageError, throwIfAborted } from '../../core/package-error';
import { readPartBounded } from '../../pptx/appearance/read';
import { validateSlideSize } from '../../pptx/appearance/size';
import type { PptxDocument, PptxPackageHandle, PptxSlide } from '../../pptx/types';
import type { AllocationOwner, DiagnosticSink, Lease, MetadataHandle, ShowIndex, ShowSlide, SlideMetadataStore, SlideShowWorkspace, SlideTimingReader, TargetResolver } from '../types';
import { relationships, relsPart } from './deck';
import { readNotes } from './notes';
import { readLinks } from './links';
import { readMedia } from './media';
import { READER_LIMITS as L } from './limits';
interface Args {
    doc: PptxDocument;
    handle?: PptxPackageHandle;
    index: ShowIndex;
    workspace: SlideShowWorkspace;
    sink: DiagnosticSink;
    timing?: SlideTimingReader;
    resolver: (slide: PptxSlide) => TargetResolver;
    signal?: AbortSignal;
    /** Injectable model-cost estimator; the construction reservation remains enforced. */
    estimate?: (slide: ShowSlide) => number;
}
interface Pending {
    job: Promise<Entry>;
    waiters: number;
    cached: boolean;
    entry?: Entry;
}
interface Entry {
    slide: ShowSlide;
    lease: Lease;
    refs: number;
    cached: boolean;
}
function cost(value: unknown): number {
    if (typeof value === 'string')
        return 24 + value.length * 2;
    if (!value || typeof value !== 'object')
        return 8;
    if (Array.isArray(value))
        return 32 + value.reduce((n, x) => n + 8 + cost(x), 0);
    let n = 48;
    for (const k of Object.keys(value))
        n += 16 + k.length * 2 + cost((value as Record<string, unknown>)[k]);
    return n;
}
/** A temporary owner covers inflate chunks, concatenation, UTF-16, parsed XML
 * and parser scratch before construction. Text-derived markup cost is reserved
 * before XML parsing. These are conservative accounting estimates, not heap telemetry. */
async function xml(handle: PptxPackageHandle, path: string, cap: number, owner: AllocationOwner, signal: AbortSignal, leases: Lease[]): Promise<XmlNode | undefined> {
    const raw = owner.lease(0, `xml:${path}`);
    if (!raw)
        throw new PackageError('resource-limit', 'XML workspace');
    leases.push(raw);
    let bytes = 0;
    const chunks = await readPartBounded(handle.pkg, path, cap, { signal, output: 'chunks', admit: n => { if (!raw.grow(n * 8 + 128))
            return false; bytes += n; return true; } });
    throwIfAborted(signal);
    if (!chunks)
        return undefined;
    const joined = new Uint8Array(bytes);
    let off = 0;
    for (const c of chunks) {
        joined.set(c, off);
        off += c.length;
    }
    const text = new TextDecoder().decode(joined);
    // Per '<' and attribute '=' upper estimate covers parser nodes, ordered child
    // arrays, namespace bindings and duplicate normalized objects. Comments may
    // overestimate; never infer allocation from untrusted ZIP size.
    let markup = 0;
    for (let i = 0; i < text.length; i++)
        if (text[i] === '<' || text[i] === '=')
            markup++;
    if (!raw.grow(markup * 1024 + 1024))
        throw new PackageError('resource-limit', 'XML object workspace');
    return parseXmlOrdered(text);
}
export function createSlideMetadataStore(args: Args): SlideMetadataStore {
    const { doc, handle, index, workspace, sink } = args;
    const cache = new Map<number, Entry>(), pending = new Map<number, Pending>(), abort = new AbortController();
    let total = 0, closed = false;
    const stop = () => store.close();
    args.signal?.addEventListener('abort', stop, { once: true });
    const diagnose = (i: number, code: string, message: string) => sink.add({ code, slideIndex: i, message });
    function dispose(e: Entry) { if (!e.cached && e.refs === 0) {
        total -= e.lease.bytes;
        e.lease.release();
    } }
    function evict(i: number) { const p = pending.get(i); if (p) {
        p.cached = false;
        pending.delete(i);
    } const e = cache.get(i); if (e) {
        cache.delete(i);
        e.cached = false;
        dispose(e);
    } }
    function room(bytes: number, except?: Entry) { for (const [i, e] of cache) {
        if (total + bytes <= L.metadataBytes)
            break;
        if (e !== except && e.refs === 0 && !pending.has(i))
            evict(i);
    } return total + bytes <= L.metadataBytes; }
    const empty = (i: number, why: string): ShowSlide => ({ index: i, hidden: index.slides[i]?.hidden ?? false, timing: { refused: why }, notes: [], links: [], media: [], builds: [] });
    function reserve(bytes: number, i: number): Lease | undefined { if (!Number.isSafeInteger(bytes) || bytes < 0 || !room(bytes))
        return undefined; const l = workspace.lease(bytes, `meta:${i}`); if (l)
        total += bytes; return l; }
    async function load(i: number): Promise<Entry> {
        const signal = abort.signal, temporary: Lease[] = [];
        let slide: ShowSlide, model: Lease | undefined;
        throwIfAborted(signal);
        const shell = workspace.lease(2048, `metadata-shell:${i}`);
        if (!shell)
            throw new PackageError('resource-limit', 'Metadata shell workspace');
        temporary.push(shell);
        try {
            if (!handle || !index.slides[i]?.partPath) {
                sink.add({ code: 'no-package', message: 'No package; slides are static' });
                slide = empty(i, 'no-package');
            }
            else if (!validateSlideSize(doc.slides[i].widthEmu, doc.slides[i].heightEmu).ok) {
                diagnose(i, 'slide-size', 'Invalid slide size');
                slide = empty(i, 'slide-size');
            }
            else {
                const part = index.slides[i].partPath!;
                const source = await xml(handle, part, L.slideBytes, workspace, signal, temporary);
                if (!source)
                    throw new PackageError('parse-error', 'Missing slide');
                const rels = relationships(await xml(handle, relsPart(part), L.presPropsBytes, workspace, signal, temporary), part);
                const note = rels.find(r => !r.external && r.type.endsWith('/notesSlide'));
                let notes: XmlNode | undefined, notesRefused = false;
                if (note)
                    try {
                        notes = await xml(handle, note.target, L.notesBytes, workspace, signal, temporary);
                    }
                    catch (e) {
                        throwIfAborted(signal);
                        notesRefused = true;
                        diagnose(i, 'notes-limit', 'Notes part refused');
                    }
                const types = rels.some(r => /\/(audio|video|media)$/.test(r.type)) ? await xml(handle, '[Content_Types].xml', L.presPropsBytes, workspace, signal, temporary) : undefined;
                // Input-derived construction estimate: XML leases already price every
                // character and markup/attribute node conservatively. Twice that cost
                // covers normalized metadata and the reader's bounded timing/build model.
                // Admit against BOTH budgets before the synchronous reader allocates.
                const bound = Math.min(L.metadataBytes, temporary.reduce((n, l) => n + l.bytes, 0) * 2 + 2048);
                model = reserve(bound, i);
                if (!model)
                    throw new PackageError('resource-limit', 'Metadata construction admission');
                const paragraphs = readNotes(notes, i, {add(d) { if(d.code === 'notes-truncated') notesRefused = true; sink.add(d); }, list: () => sink.list()});
                const timing = args.timing?.(source, i, args.resolver(doc.slides[i]), sink);
                slide = { index: i, hidden: index.slides[i].hidden, ...timing, timing: notesRefused ? { refused: 'notes-limit' } : timing?.timing, notes: paragraphs, links: readLinks(source, rels, handle.slideParts, i, sink), media: readMedia(rels, types, handle.pkg, i, sink), builds: timing?.builds ?? [] };
                throwIfAborted(signal);
            }
            const bytes = args.estimate?.(slide) ?? cost(slide);
            let lease = model;
            if (lease) {
                if (!Number.isSafeInteger(bytes) || bytes < 0)
                    throw new PackageError('resource-limit', 'Invalid metadata estimate');
                if (bytes > lease.bytes) {
                    const growth = bytes - lease.bytes;
                    if (!room(growth) || !lease.grow(growth)) {
                        total -= lease.bytes;
                        lease.release();
                        model = undefined;
                        lease = undefined;
                    }
                    else
                        total += growth;
                }
                else {
                    const refund = lease.bytes - bytes;
                    lease.shrink(refund);
                    total -= refund;
                }
            }
            else
                lease = reserve(bytes, i);
            if (!lease) {
                diagnose(i, 'metadata-limit', 'Aggregate metadata admission refused; slide is static');
                slide = empty(i, 'metadata-limit');
                const fallback = reserve(cost(slide), i);
                if (!fallback)
                    throw new PackageError('resource-limit', 'Static metadata admission refused');
                return { slide, lease: fallback, refs: 0, cached: true };
            }
            model = undefined;
            return { slide, lease, refs: 0, cached: true };
        }
        catch (e) {
            if (model) {
                total -= model.bytes;
                model.release();
                model = undefined;
            }
            throwIfAborted(signal);
            diagnose(i, 'slide-refused', e instanceof Error ? e.message : 'Slide read refused');
            slide = empty(i, 'read-limit');
            const lease = reserve(cost(slide), i);
            if (!lease)
                throw new PackageError('resource-limit', 'Static metadata admission refused');
            return { slide, lease, refs: 0, cached: true };
        }
        finally {
            for (const lease of temporary)
                lease.release();
        }
    }
    function handleFor(e: Entry): MetadataHandle { e.refs++; let released = false; return { slide: e.slide, attach(bytes, _tag) { if (released || closed || !Number.isSafeInteger(bytes) || bytes < 0 || !room(bytes, e))
            return false; if (!e.lease.grow(bytes))
            return false; total += bytes; return true; }, release() { if (released)
            return; released = true; e.refs--; dispose(e); } }; }
    const store: SlideMetadataStore = {
        async acquire(i, signal) {
            throwIfAborted(signal);
            throwIfAborted(abort.signal);
            if (!Number.isInteger(i) || i < 0 || i >= doc.slides.length)
                throw new RangeError('slide index');
            const cached = cache.get(i);
            if (cached) {
                cache.delete(i);
                cache.set(i, cached);
                return handleFor(cached);
            }
            let p = pending.get(i);
            if (!p) {
                const record: Pending = { job: undefined as never, waiters: 0, cached: true };
                record.job = load(i).then(e => {
                    // The job owns one reference until all waiting acquisitions settle.
                    e.refs++;
                    record.entry = e;
                    e.cached = record.cached && !closed;
                    if (e.cached)
                        cache.set(i, e);
                    return e;
                });
                p = record;
                pending.set(i, p);
            }
            p.waiters++;
            try {
                const entry = await p.job;
                throwIfAborted(signal);
                throwIfAborted(abort.signal);
                return handleFor(entry);
            }
            finally {
                p.waiters--;
                if (p.waiters === 0) {
                    if (pending.get(i) === p)
                        pending.delete(i);
                    if (p.entry) {
                        p.entry.refs--;
                        dispose(p.entry);
                    }
                }
            }
        }, evict, resident: () => [...cache.keys()], close() { if (closed)
            return; closed = true; abort.abort(new PackageError('cancelled', 'Store closed')); args.signal?.removeEventListener('abort', stop); for (const i of cache.keys())
            evict(i); }
    };
    if (args.signal?.aborted)
        store.close();
    return store;
}
