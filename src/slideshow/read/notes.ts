import { attrs, getChild, getChildren, elementChildren, textOf, type XmlNode } from '../../core/xml';
import type { DiagnosticSink, NoteParagraph } from '../types';
import { READER_LIMITS } from './limits';
/** Body placeholders only: date, slide number and thumbnail are not speaker notes. */
export function readNotes(xml: XmlNode | undefined, index: number, sink: DiagnosticSink): NoteParagraph[] {
    const result: NoteParagraph[] = [];
    let left = READER_LIMITS.notesChars, truncated = false;
    for (const sp of getChildren(getChild(getChild(xml, 'cSld'), 'spTree'), 'sp')) {
        if (attrs(getChild(getChild(getChild(sp, 'nvSpPr'), 'nvPr'), 'ph')).type !== 'body')
            continue;
        for (const p of getChildren(getChild(sp, 'txBody'), 'p')) {
            let text = '';
            for (const [k, n] of elementChildren(p)) {
                const value = k === 'r' || k === 'fld' ? textOf(getChild(n, 't')) : k === 'br' ? '\n' : '';
                if (value.length > left)
                    truncated = true;
                let take = Math.min(left, value.length);
                // Never split a surrogate pair at the text cap.
                if (take < value.length && take > 0 && /^[\uDC00-\uDFFF]/.test(value.slice(take, take + 1)) && /[\uD800-\uDBFF]$/.test(value.slice(0, take)))
                    take--;
                text += value.slice(0, take);
                left -= take;
            }
            const pr = getChild(p, 'pPr'), level = Number(attrs(pr).lvl ?? 0), char = attrs(getChild(pr, 'buChar')).char, auto = attrs(getChild(pr, 'buAutoNum')).type;
            if (left > 0 || text)
                result.push({ level: Number.isInteger(level) && level >= 0 && level <= 8 ? level : 0, ...(char || auto ? { bullet: char ?? auto } : {}), text });
        }
    }
    if (truncated)
        sink.add({ code: 'notes-truncated', slideIndex: index, message: 'Notes exceed 65536 characters; truncated' });
    return result;
}
