/** Spec 6.1 and 7. Byte limits are inflated bytes, not ZIP declarations. */
export const READER_LIMITS = Object.freeze({
  /** Inflated XML per slide part. */
  slideBytes: 4 * 1024 * 1024,
  notesBytes: 1024 * 1024,
  presPropsBytes: 1024 * 1024,
  /** Not in spec 6.1: caps for the small structural parts the reader needs (rels, content types, presentation). */
  relsBytes: 1024 * 1024,
  contentTypesBytes: 1024 * 1024,
  presentationBytes: 4 * 1024 * 1024,
  timeNodes: 2000,
  nesting: 64,
  conditions: 10000,
  keyframesPerBehaviour: 1000,
  keyframesPerSlide: 20000,
  motionPathBytes: 65536,
  motionPathCommands: 10000,
  notesChars: 65536,
  metadataBytes: 64 * 1024 * 1024,
  mediaBytes: 512 * 1024 * 1024,
  formulaBytes: 1024,
  formulaTokens: 256,
  formulaDepth: 32,
})
