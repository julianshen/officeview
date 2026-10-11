/** Typed failures of the bounded large-source readers (zip-stream, inflater,
 * xml-stream). src/core must not import src/xlsx, so the codes are a local
 * subset; tests pin that each one is a member of SOURCE_ERROR_CODES. */
export type PackageErrorCode = 'cancelled' | 'resource-limit' | 'unsupported' | 'parse-error' | 'storage-error' | 'quota'

export class PackageError extends Error {
  readonly code: PackageErrorCode
  readonly partId?: string

  constructor(code: PackageErrorCode, message: string, partId?: string) {
    super(partId === undefined ? message : `${message} (${partId})`)
    this.name = 'PackageError'
    this.code = code
    if (partId !== undefined) this.partId = partId
  }
}

export function isPackageError(error: unknown): error is PackageError {
  return error instanceof PackageError
}

export function throwIfAborted(signal: AbortSignal | undefined, partId?: string): void {
  if (signal?.aborted) throw new PackageError('cancelled', 'operation cancelled', partId)
}
