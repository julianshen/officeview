import type JSZip from 'jszip';
import { OfficePackage } from '../../core/zip';
import { attrs, getChild, getChildren, parseXml, type XmlNode } from '../../core/xml';
import { throwIfAborted } from '../../core/package-error';
import { readXmlBounded, readPartBounded } from '../../pptx/appearance/read';
import { validateSlideSize } from '../../pptx/appearance/size';
import type { PptxDocument, PptxPackageHandle } from '../../pptx/types';
import type { DiagnosticSink, ShowIndex, ShowSettings } from '../types';
import { READER_LIMITS } from './limits';
export interface Relationship {
    id: string;
    type: string;
    target: string;
    external: boolean;
}
export function resolvePart(part: string, target: string): string {
    const stack = target.startsWith('/') ? [] : part.split('/').slice(0, -1);
    for (const p of target.split('/')) {
        if (p === '..') {
            if (!stack.length)
                return '';
            stack.pop();
        }
        else if (p && p !== '.')
            stack.push(p);
    }
    return stack.join('/');
}
export function relsPart(part: string): string { const n = part.lastIndexOf('/'); return `${part.slice(0, n + 1)}_rels/${part.slice(n + 1)}.rels`; }
export function relationships(xml: XmlNode | undefined, part: string): Relationship[] {
    return getChildren(xml, 'Relationship').map(n => { const a = attrs(n); const external = a.TargetMode === 'External'; return { id: a.Id, type: a.Type, target: external ? a.Target : resolvePart(part, a.Target ?? ''), external }; });
}
export async function readRelationships(handle: PptxPackageHandle, part: string, signal?: AbortSignal): Promise<Relationship[]> {
    return relationships(await readXmlBounded(handle.pkg, relsPart(part), READER_LIMITS.presPropsBytes, signal), part);
}
const defaults = (): ShowSettings => ({ loop: false, showNarration: false, showAnimation: true, useTimings: true, type: 'present', kioskRestartMs: 300000 });
/** Re-readable sources only. Presentation order comes from sldIdLst, never filenames. */
export async function openShowPackage(source: Blob | ArrayBuffer | Uint8Array | string): Promise<PptxPackageHandle> {
    const data = typeof source === 'string' ? new Uint8Array(await (await fetch(source)).arrayBuffer()) : source instanceof Blob ? new Uint8Array(await source.arrayBuffer()) : source;
    const pkg = await OfficePackage.load(data);
    const roots = relationships(await readXmlBounded(pkg, '_rels/.rels', READER_LIMITS.presPropsBytes), '');
    const part = roots.find(r => !r.external && r.type?.endsWith('/officeDocument'))?.target ?? 'ppt/presentation.xml';
    const handle = { pkg, slideParts: [] as string[] };
    const rels = await readRelationships(handle, part);
    const presentation = await readXmlBounded(pkg, part, READER_LIMITS.slideBytes);
    for (const id of getChildren(getChild(presentation, 'sldIdLst'), 'sldId')) {
        const r = rels.find(r => r.id === attrs(id).id && !r.external && r.type?.endsWith('/slide'));
        if (r)
            handle.slideParts.push(r.target);
    }
    return handle;
}
/** Read only through the root opening tag. The bounded reader aborts the inflate
 * at the first data chunk containing that tag; no slide body is parsed/cached. */
async function slideRoot(handle: PptxPackageHandle, path: string, signal?: AbortSignal): Promise<XmlNode | undefined> {
    const file = handle.pkg.entry(path);
    if (!file)
        return undefined;
    let prefix = '', root: XmlNode | undefined;
    const done = new Error('root-complete');
    // Adapt only this stream, not the package/file shared by another consumer.
    const source = { entry: () => ({ dir: false, internalStream: () => {
                const stream = (file as unknown as {
                    internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
                }).internalStream('uint8array');
                const on = stream.on.bind(stream);
                const decoder = new TextDecoder();
                let size = 0;
                stream.on = ((event: string, callback: (...args: any[]) => void) => on(event as never, event === 'data' ? ((bytes: Uint8Array) => {
                    throwIfAborted(signal, path);
                    size += bytes.length;
                    if (size > READER_LIMITS.slideBytes)
                        throw new Error('slide root limit');
                    prefix += decoder.decode(bytes, { stream: true });
                    // XML declarations/comments may precede the root. Do not match inside them.
                    const clean = prefix.replace(/^\uFEFF/, '').replace(/^(?:\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->)+/, '');
                    const m = /^<([\w.-]+:)?sld\b(?:"[^"]*"|'[^']*'|[^'">])*>/.exec(clean);
                    if (m) {
                        root = parseXml(m[0].replace(/\/?\s*>$/, '/>'));
                        throw done;
                    }
                    callback(bytes);
                }) as never : callback as never)) as typeof stream.on;
                return stream;
            } }) as unknown as ReturnType<OfficePackage['entry']> };
    try {
        await readPartBounded(source, path, READER_LIMITS.slideBytes, { signal, output: 'chunks' });
    }
    catch (e) {
        if (e !== done)
            throw e;
    }
    return root;
}
export async function readShowIndex(doc: PptxDocument, handle: PptxPackageHandle | undefined, sink: DiagnosticSink, signal?: AbortSignal): Promise<ShowIndex> {
    throwIfAborted(signal);
    const settings = defaults();
    const slides: {
        index: number;
        hidden: boolean;
        partPath?: string;
    }[] = [];
    const diagnose = (code: string, message: string, slideIndex?: number) => sink.add({ code, message, slideIndex });
    if (!validateSlideSize(doc.slideWidthEmu, doc.slideHeightEmu).ok)
        diagnose('slide-size', 'Invalid slide size; show uses static slides');
    if (!handle)
        diagnose('no-package', 'No package; slides are static');
    else {
        try {
            const packageRels = relationships(await readXmlBounded(handle.pkg, '_rels/.rels', READER_LIMITS.presPropsBytes, signal), '');
            const presentation = packageRels.find(r => !r.external && r.type?.endsWith('/officeDocument'))?.target ?? 'ppt/presentation.xml';
            const rels = await readRelationships(handle, presentation, signal);
            const part = rels.find(r => !r.external && r.type?.endsWith('/presProps'))?.target;
            const pr = part ? getChild(await readXmlBounded(handle.pkg, part, READER_LIMITS.presPropsBytes, signal), 'showPr') : undefined;
            const a = attrs(pr);
            for (const key of ['loop', 'showNarration', 'showAnimation', 'useTimings'] as const)
                if (a[key] !== undefined) {
                    if (['true', '1', 'false', '0'].includes(a[key]))
                        settings[key] = ['true', '1'].includes(a[key]);
                    else
                        diagnose('show-properties', `Invalid ${key}`);
                }
            if (getChildren(pr, 'kiosk').length) {
                settings.type = 'kiosk';
                const v = attrs(getChild(pr, 'kiosk')).restart;
                if (v !== undefined) {
                    const n = Number(v);
                    if (/^\d+$/.test(v) && n <= 0xffffffff)
                        settings.kioskRestartMs = n;
                    else
                        diagnose('show-properties', 'Invalid kiosk restart');
                }
            }
            if (getChildren(pr, 'browse').length) {
                settings.type = 'present';
                diagnose('browse', 'Browse uses present mode');
            }
            if (getChildren(pr, 'custShow').length)
                diagnose('custom-show', 'Custom show uses full deck');
            else if (!getChildren(pr, 'browse').length) {
                const r = attrs(getChild(pr, 'sldRg'));
                if (r.st !== undefined || r.end !== undefined) {
                    const st = Number(r.st), end = Number(r.end);
                    if (Number.isInteger(st) && Number.isInteger(end) && st >= 1 && end >= st && end <= doc.slides.length)
                        settings.range = { st: st - 1, end: end - 1 };
                    else
                        diagnose('slide-range', 'Invalid range; full deck used');
                }
            }
        }
        catch (e) {
            throwIfAborted(signal);
            Object.assign(settings, defaults());
            diagnose('show-properties', 'Show properties refused; defaults used');
        }
    }
    for (let index = 0; index < doc.slides.length; index++) {
        throwIfAborted(signal);
        const partPath = handle?.slideParts[index];
        let hidden = false;
        if (handle && partPath)
            try {
                const a = attrs(await slideRoot(handle, partPath, signal));
                hidden = a.show === '0' || a.show === 'false';
            }
            catch {
                throwIfAborted(signal);
                diagnose('slide-root', 'Slide root refused', index);
            }
        slides.push({ index, hidden, ...(partPath ? { partPath } : {}) });
    }
    return { slides, settings, diagnostics: [...sink.list()] };
}
