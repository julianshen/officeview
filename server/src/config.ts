/** Server configuration (spec §9.1 `config.ts`): keys and flags, fail-closed. */
import { loadKeyring, type Keyring } from './keyring'

export class ConfigError extends Error {
  constructor(readonly code: 'missing-key' | 'bad-key' | 'bad-arg' | 'missing-audit', message: string) {
    super(message)
    this.name = 'ConfigError'
  }
}

export interface ServerConfig {
  port: number
  demo: boolean
  docDir: string
  originAllowlist: string[]
  keyring: Keyring
  /** Durable audit store path (`OV_AUDIT_DB`); `':memory:'` under `--demo`. */
  auditDb: string
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

export function loadConfig(env: Record<string, string | undefined>, argv: readonly string[]): ServerConfig {
  const demo = argv.includes('--demo')
  const portArg = flag(argv, '--port') ?? env.PORT ?? '8787'
  const port = Number(portArg)
  if (!Number.isInteger(port) || port < 0 || port > 65535) throw new ConfigError('bad-arg', `bad port ${portArg}`)
  const origins = [...(flag(argv, '--origin')?.split(',') ?? []), ...(env.OV_STREAM_ORIGINS?.split(',') ?? [])].map((s) => s.trim()).filter(Boolean)
  const docDir = flag(argv, '--doc-dir') ?? env.OV_STREAM_DOC_DIR ?? ''
  const { keyring, ephemeral } = loadKeyring(env, demo)
  const auditDb = demo ? ':memory:' : env.OV_AUDIT_DB ?? ''
  if (!auditDb) throw new ConfigError('missing-audit', 'OV_AUDIT_DB=<path> is required outside --demo')
  return { port, demo, docDir, originAllowlist: origins, keyring, auditDb, ephemeral }
}
