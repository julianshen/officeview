/** Demo stream server entry (grows in Stage 1c Task 9). */
import { loadConfig } from './config'

if (import.meta.main) {
  const config = loadConfig(process.env, process.argv.slice(2))
  if (config.ephemeral) console.warn('EPHEMERAL KEY — sessions die with this process')
}
