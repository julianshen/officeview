import { attrs, getChildren, type XmlNode } from '../../core/xml';
import type { OfficePackage } from '../../core/zip';
import type { DiagnosticSink, MediaRecord } from '../types';
import type { Relationship } from './deck';
import { READER_LIMITS } from './limits';
export function readMedia(rels: readonly Relationship[], types: XmlNode | undefined, pkg: OfficePackage, index: number, sink: DiagnosticSink): MediaRecord[] {
    const out: MediaRecord[] = [];
    for (const r of rels) {
        if (!/\/(audio|video|media)$/.test(r.type))
            continue;
        if (r.external) {
            sink.add({ code: 'media-refused', slideIndex: index, effect: r.id, message: 'External media is not read' });
            continue;
        }
        const entry = pkg.entry(r.target), size = (entry as unknown as {
            _data?: {
                uncompressedSize?: number;
            };
        })?._data?.uncompressedSize;
        if (!entry || !Number.isSafeInteger(size) || size! < 0) {
            sink.add({ code: 'media-refused', slideIndex: index, effect: r.id, message: 'Missing media or invalid declared size' });
            continue;
        }
        const override = getChildren(types, 'Override').find(n => attrs(n).PartName === '/' + r.target);
        const def = getChildren(types, 'Default').find(n => attrs(n).Extension === r.target.split('.').pop());
        if (size! > READER_LIMITS.mediaBytes)
            sink.add({ code: 'media-limit', slideIndex: index, effect: r.id, message: 'Declared media exceeds 512 MiB; playback must refuse' });
        out.push({ rId: r.id, partPath: r.target, contentType: attrs(override ?? def).ContentType ?? 'application/octet-stream', declaredBytes: size! });
    }
    return out;
}
