/** Debug-only (see cfEvidence.ts): installs the Stage 5 evidence inside the e2e viewer worker when its name carries `cfev`. */
import { installCfEvidence } from './cfEvidence'
import { workerNameSegments } from './forcePaged'

if (workerNameSegments((globalThis as { name?: string }).name).includes('cfev')) installCfEvidence()
