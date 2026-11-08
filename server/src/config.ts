/** Server configuration (spec §9.1 `config.ts`): keys and flags, fail-closed. */

export class ConfigError extends Error {
  constructor(readonly code: 'missing-key' | 'bad-key' | 'bad-arg', message: string) {
    super(message)
    this.name = 'ConfigError'
  }
}

export interface ServerConfig {
  port: number
  demo: boolean
  docDir: string
  originAllowlist: string[]
  masterKey: Uint8Array
  /** True when `--demo` generated the key: sessions die with the process. */
  ephemeral: boolean
}

function flag(argv: readonly string[], name: string): string | undefined {
  const i = argv.indexOf(name)
  if (i === -1) return undefined
  const v = argv[i + 1]
  if (v === undefined || v.startsWith('--')) throw new ConfigError('bad-arg', `${name} needs a value`)
  return v
}

function decodeKey(b64: string): Uint8Array {
  let bytes: Uint8Array
  try { bytes = Uint8Array.from(atob(b64.trim()), (c) => c.charCodeAt(0)) } catch { throw new ConfigError('bad-key', 'OV_STREAM_KEY is not base64') }
  if (bytes.length !== 32) throw new ConfigError('bad-key', `OV_STREAM_KEY must decode to 32 bytes, got ${bytes.length}`)
  return bytes
}

export function loadConfig(env: Record<string, string | undefined>, argv: readonly string[]): ServerConfig {
  const demo = argv.includes('--demo')
  const portArg = flag(argv, '--port') ?? env.PORT ?? '8787'
  const port = Number(portArg)
  if (!Number.isInteger(port) || port < 0 || port > 65535) throw new ConfigError('bad-arg', `bad port ${portArg}`)
  const origins = [...(flag(argv, '--origin')?.split(',') ?? []), ...(env.OV_STREAM_ORIGINS?.split(',') ?? [])].map((s) => s.trim()).filter(Boolean)
  const docDir = flag(argv, '--doc-dir') ?? env.OV_STREAM_DOC_DIR ?? ''
  const raw = env.OV_STREAM_KEY
  let masterKey: Uint8Array
  let ephemeral = false
  if (raw) masterKey = decodeKey(raw)
  else if (demo) { masterKey = crypto.getRandomValues(new Uint8Array(32)); ephemeral = true }
  else throw new ConfigError('missing-key', 'OV_STREAM_KEY is required outside --demo')
  return { port, demo, docDir, originAllowlist: origins, masterKey, ephemeral }
}
