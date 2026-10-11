/** Page <-> kernel-worker messages for the S8.4 scenarios (types only). */
import type { KernelSample, QuerySample, ScenarioInfo } from './kernelScenarios'

export type KernelRequest =
  | { type: 'setup'; id: number; scenario: string; rows: number | null }
  | { type: 'kernel'; id: number }
  | { type: 'query'; id: number }

export type KernelReply =
  | { type: 'ready'; id: number; info: ScenarioInfo; prep: { metrics: Record<string, number>; notes: string[] } }
  | { type: 'kernel'; id: number; sample: KernelSample }
  | { type: 'query'; id: number; sample: QuerySample }
  | { type: 'error'; id: number; message: string }
