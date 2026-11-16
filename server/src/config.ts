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
  /**
   * Piece 4 (spec §7, owner decision 1): `strictMarks` (`--strict-marks` or `OV_STRICT_MARKS=1`) plays every moving
   * step `composited`, so the visible mark never leaves its page anchor. Off by default (the mark moves with a mover).
   */
  animation: { strictMarks: boolean }
}

/**
 * TEST-ONLY server options (the parity harness, piece 4 Task 14). Set only by the code that builds the server; never
 * read from the config file, the environment or an `AuthDecision` (an `authorize` hook cannot turn Layer-1 off).
 * `forensicOff`: show assets ship without Layer-1 (audited `unmarked`, reason `forensic-off`).
 */
export interface TestOnlyOptions { forensicOff?: boolean }

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
  const strict = env.OV_STRICT_MARKS
  if (strict !== undefined && strict !== '0' && strict !== '1') throw new ConfigError('bad-arg', 'OV_STRICT_MARKS must be 0 or 1')
  const animation = { strictMarks: argv.includes('--strict-marks') || strict === '1' }
  return { port, demo, docDir, originAllowlist: origins, keyring, auditDb, ephemeral, animation }
}
