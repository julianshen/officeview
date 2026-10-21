import { attrs, getChild, getChildren, elementChildren, type XmlNode } from '../../core/xml';
import type { DiagnosticSink, LinkSpec } from '../types';
import type { Relationship } from './deck';
const jumps = new Set(['nextslide', 'previousslide', 'firstslide', 'lastslide', 'endshow']);
/** MS-OE376 2.1.1315. Unsupported custom shows/lastslideviewed/macros are explicit refusals. */
export function readLinks(xml: XmlNode, rels: readonly Relationship[], parts: readonly string[], index: number, sink: DiagnosticSink): LinkSpec[] {
    const links: LinkSpec[] = [];
    function action(n: XmlNode): LinkSpec['action'] {
        const a = attrs(n), r = rels.find(r => r.id === a.id), verb = a.action ?? '';
        if (verb.startsWith('ppaction://hlinkshowjump?jump=')) {
            const v = verb.slice('ppaction://hlinkshowjump?jump='.length);
            if (jumps.has(v))
                return { kind: 'jump', verb: v as 'nextslide' };
        }
        if (verb === 'ppaction://hlinksldjump' || !verb) {
            if (r && !r.external && r.type.endsWith('/slide')) {
                const i = parts.indexOf(r.target);
                if (i >= 0)
                    return { kind: 'slide', index: i };
            }
        }
        if (!verb && r?.external) {
            try {
                const u = new URL(r.target);
                if (['http:', 'https:', 'mailto:'].includes(u.protocol))
                    return { kind: 'url', href: r.target };
            }
            catch { }
        }
        return { kind: 'refused', reason: verb || 'unsafe-or-missing-target' };
    }
    function add(n: XmlNode, spid: string, run?: {
        p: number;
        r: number;
    }) { const a = action(n); links.push({ spid, ...(run ? { run } : {}), action: a }); if (a.kind === 'refused')
        sink.add({ code: 'link-refused', slideIndex: index, effect: `${spid}:${run?.p ?? ''}:${run?.r ?? ''}`, message: a.reason }); }
    function walk(n: XmlNode) {
        for (const [kind, shape] of elementChildren(n)) {
            if (kind === 'grpSp') {
                walk(shape);
                continue;
            }
            const nv = getChild(shape, 'nvSpPr') ?? getChild(shape, 'nvPicPr') ?? getChild(shape, 'nvGraphicFramePr') ?? getChild(shape, 'nvCxnSpPr'), cn = getChild(nv, 'cNvPr'), spid = attrs(cn).id;
            if (!spid)
                continue;
            for (const link of getChildren(cn, 'hlinkClick'))
                add(link, spid);
            getChildren(getChild(shape, 'txBody'), 'p').forEach((p, pi) => {
                let ri = 0;
                for (const [k, r] of elementChildren(p))
                    if (k === 'r' || k === 'fld' || k === 'br') {
                        for (const link of getChildren(getChild(r, 'rPr'), 'hlinkClick'))
                            add(link, spid, { p: pi, r: ri });
                        ri++;
                    }
            });
        }
    }
    walk(getChild(getChild(xml, 'cSld'), 'spTree') ?? {});
    return links;
}
