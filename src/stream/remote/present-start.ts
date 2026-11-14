/**
 * The streamed deck's part of `present()` (design §4.3; parent §8.2, §12a.7).
 * Imported statically by the lazy remote view only, so none of it is eager:
 * the eager `startPresent` (components/present.ts) runs the shared gesture
 * prologue over the launch prepared here.
 *
 * Synchronous checks, from the manifest as it is now: presenting allowed by
 * the host and the server (`can.present`), at least one shown slide, a
 * numeric `from` inside the shown slides (a streamed index: hidden units are
 * not counted, R19), and `presenterView: true` only with `can.notes`. Without
 * `can.notes` no presenter window is opened for 'auto' either, so a
 * non-presenter never gets a blank `window.open` (§8.2). The show itself
 * (show.ts) loads lazily, after the surface is up.
 */
import { SlideShowError, startPresent, type PresentRequest, type ShowLaunch, type SlideShowCallbacks, type SlideShowErrorCode, type SlideShowSession } from '../../components/present'
import type { RemoteOfficeDoc } from './remote-doc'

/** @internal Tests defer or fail the remote show module through this. */
export const remoteShowLoader = { load: () => import('./show') }

/** The synchronous document checks (before the shared prologue's activation check, as the local order). */
export function remoteRefusal(doc: RemoteOfficeDoc, opts: PresentRequest['opts'] = {}): SlideShowErrorCode | 0 {
  const m = doc.manifest
  const shown = m.units.list.filter((u) => !u.hidden).length
  const from = opts.from ?? 'beginning'
  return !m.can.present || !shown || opts.presenterView === true && !m.can.notes ? 'not-presentable'
    : typeof from == 'number' && !(Number.isInteger(from) && from >= 0 && from < shown) ? 'invalid-slide' : 0
}

/**
 * `startPresent` for a streamed deck: call synchronously inside the user
 * gesture. The host's `allow` is checked first (by the shared prologue's
 * rule, here so a refused start sends nothing), then the document checks.
 * Without `can.notes` the prologue is told `presenterView: false`, so 'auto'
 * opens no presenter window (true was refused above).
 */
export function presentRemote(req: Omit<PresentRequest, 'doc' | 'launch'> & { doc: RemoteOfficeDoc }): Promise<SlideShowSession> {
  const { doc } = req
  const code = req.allow ? remoteRefusal(doc, req.opts) : 'not-presentable'
  if (code) return Promise.reject(new SlideShowError(code))
  const launch: ShowLaunch = () => remoteShowLoader.load().then((x) => (s) =>
    // The remote view supplies every navigation callback (unit indices).
    x.runRemoteShow({ ...s, doc, callbacks: s.callbacks as SlideShowCallbacks }) as Promise<SlideShowSession>)
  const opts = doc.manifest.can.notes ? req.opts : { ...req.opts, presenterView: false }
  return startPresent({ ...req, opts, launch })
}
