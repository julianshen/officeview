/**
 * Dev-server worker entry for the e2e page: the shipped dist/worker.js source
 * (src/worker/worker-entry.ts) loaded from inside src/debug, so the harness
 * server's locked-down fs.allow serves it through the module graph.
 */
// Capability denial (S10.3) runs before the entry module evaluates.
import './workerCaps'
import '../worker/worker-entry'
// The kernel imports its large-source opener on demand; named here for the same reason.
import '../worker/viewerLargeSource'
