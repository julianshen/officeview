/** Typed failures of the bounded large-source readers (zip-stream, inflater,
 * xml-stream). src/core must not import src/xlsx, so the codes are a local
 * subset; tests pin that each one is a member of SOURCE_ERROR_CODES. */
export type PackageErrorCode = 'cancelled' | 'resource-limit' | 'unsupported' | 'parse-error' | 'storage-error' | 'quota'

/** Machine-readable cause for the failures callers route on (messages are for people). */
export type PackageErrorReason =
  | 'entry-limit' | 'central-directory-limit' | 'name-limit' | 'name-encoding' | 'name-invalid'
  | 'duplicate-name' | 'case-duplicate-name' | 'encrypted' | 'method' | 'multi-disk'
  | 'size-mismatch' | 'truncated-deflate'
  /** The worker memory pool cannot hold an import's construction peak (S13.4, spec §12.4). */
  | 'memory-budget'

export class PackageError extends Error {
  readonly code: PackageErrorCode
  readonly partId?: string
  readonly reason?: PackageErrorReason

  constructor(code: PackageErrorCode, message: string, partId?: string, reason?: PackageErrorReason) {
    super(partId === undefined ? message : `${message} (${partId})`)
    this.name = 'PackageError'
    this.code = code
    if (partId !== undefined) this.partId = partId
    if (reason !== undefined) this.reason = reason
  }
}

export function isPackageError(error: unknown): error is PackageError {
  return error instanceof PackageError
}

export function throwIfAborted(signal: AbortSignal | undefined, partId?: string): void {
  if (signal?.aborted) throw new PackageError('cancelled', 'operation cancelled', partId)
}
