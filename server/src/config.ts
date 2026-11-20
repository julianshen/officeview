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
  /**
   * Piece 5 (spec §4.2): `OV_MEDIA` (`on` | `off`, default `on`), `OV_FFMPEG` (default `ffmpeg` on PATH) and
   * `OV_FFPROBE` (default the `ffprobe` next to the resolved ffmpeg). Passed to `probeFfmpeg` at startup; a failed
   * probe or `off` means `can.media: false`.
   */
  media: { OV_MEDIA: 'on' | 'off'; OV_FFMPEG?: string; OV_FFPROBE?: string }
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
  const mediaMode = env.OV_MEDIA ?? 'on'
  if (mediaMode !== 'on' && mediaMode !== 'off') throw new ConfigError('bad-arg', 'OV_MEDIA must be on or off')
  const media: ServerConfig['media'] = { OV_MEDIA: mediaMode }
  if (env.OV_FFMPEG) media.OV_FFMPEG = env.OV_FFMPEG
  if (env.OV_FFPROBE) media.OV_FFPROBE = env.OV_FFPROBE
  return { port, demo, docDir, originAllowlist: origins, keyring, auditDb, ephemeral, animation, media }
}
