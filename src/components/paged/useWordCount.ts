/**
 * The document's word count, counted in idle slices after render (spec
 * §6.11). null while counting, and for PPTX and static XLSX, which show no
 * count. One AbortController per document; a new document aborts the old run.
 */
import { useEffect, useState } from 'react'
import type { DocxDocument } from '../../docx/types'
import type { OfficeDocSource } from '../OfficeDoc'
import { countDocumentWords } from './wordCount'

export function useWordCount(document: OfficeDocSource): number | null {
  const [state, setState] = useState<{ doc: OfficeDocSource; words: number } | null>(null)
  useEffect(() => {
    if (!('sections' in document)) return
    const controller = new AbortController()
    countDocumentWords(document as DocxDocument, { signal: controller.signal })
      .then((words) => setState({ doc: document, words }), () => { /* superseded */ })
    return () => controller.abort()
  }, [document])
  // a stale count never shows for a replacement document
  return state && state.doc === document ? state.words : null
}
