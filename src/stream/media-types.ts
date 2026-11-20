/**
 * Piece 5 (plan X14): the typed media refusals, shared by the server and the client. Type-only: nothing here
 * reaches the eager bundle.
 */
export type MediaRefusal = 'media-too-long' | 'media-too-large' | 'media-timeout' | 'media-busy' | 'media-unreadable' | 'media-linked' | 'media-unavailable'
