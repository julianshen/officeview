/**
 * S10.3 capability denial inside the e2e viewer worker. Playwright init
 * scripts never reach a dedicated worker, so the e2e page names the denied
 * capabilities in the Worker `name` option (`caps:<flag>,<flag>`) and this
 * module, imported before the worker entry, removes them first. An empty name
 * changes nothing. Other `;`-separated name segments (S13 `paged`) are ignored here. Flags:
 *   no-offscreen  delete OffscreenCanvas
 *   no-storage    no OPFS (getDirectory rejects) and no IndexedDB
 *   quota-zero    navigator.storage.estimate() reports no free quota
 *   idb-quota     no OPFS, and every IndexedDB put throws QuotaExceededError
 *   no-fonts      delete FontFace and self.fonts
 */
import { workerNameSegments } from './forcePaged'

const scope = globalThis as unknown as Record<string, unknown> & { name?: string; navigator?: Navigator }
const caps = workerNameSegments(scope.name).find(segment => segment.startsWith('caps:'))
const flags = new Set(caps ? caps.slice(5).split(',') : [])
const storage = scope.navigator?.storage as (StorageManager & Record<string, unknown>) | undefined
const replace = (target: object | undefined, key: string, value: unknown) => {
  if (target) Object.defineProperty(target, key, { configurable: true, value })
}
// StorageManager methods are patched on the prototype too: an engine may hand
// out a fresh navigator.storage wrapper on each access.
const replaceStorage = (key: string, value: unknown) => { replace(storage, key, value); replace(storage && Object.getPrototypeOf(storage), key, value) }
// Interface objects are own properties of the global; accessors such as
// indexedDB/fonts live on its prototype, so they are shadowed with undefined.
const remove = (key: string) => { delete scope[key]; if (scope[key] !== undefined) replace(scope, key, undefined) }
if (flags.has('no-offscreen')) remove('OffscreenCanvas')
if (flags.has('no-fonts')) { remove('FontFace'); remove('fonts') }
if (flags.has('no-storage')) {
  replaceStorage('getDirectory', () => Promise.reject(new DOMException('denied by S10.3 harness', 'SecurityError')))
  remove('indexedDB')
}
if (flags.has('idb-quota')) {
  replaceStorage('getDirectory', () => Promise.reject(new DOMException('denied by S10.3 harness', 'SecurityError')))
  replace(IDBObjectStore.prototype, 'put', () => { throw new DOMException('quota denied by S10.3 harness', 'QuotaExceededError') })
}
if (flags.has('quota-zero')) replaceStorage('estimate', async () => ({ quota: 0, usage: 0 }))
