export type FontVariant = 'regular' | 'bold' | 'italic' | 'boldItalic'
export interface EmbeddedFontFace {
  /** Requested OOXML family name, never rewritten to a registration alias. */
  family: string
  variant: FontVariant
  partPath: string
  relationshipId: string
  /** Converted and permission-checked SFNT. */
  bytes: Uint8Array
}
export interface FontDiagnostic {
  kind: 'external-font' | 'missing-font' | 'malformed-font' | 'restricted-font' | 'font-limit' | 'font-load-failed'
  message: string
  family: string
  variant: FontVariant
  relationshipId?: string
  partPath?: string
}
