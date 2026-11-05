import { afterEach, expect, it } from 'vitest'
import { ACK_ATTRIBUTES, qualifiesAck } from '../src/debug/harness/viewerAck'

afterEach(() => document.body.replaceChildren())
function observe() {
  const observer = new MutationObserver(() => {})
  observer.observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ACK_ATTRIBUTES })
  return observer
}
it('acknowledges an exiting inert dialog before its approved animation removes the DOM', () => {
  const dialog = document.createElement('div')
  dialog.role = 'dialog'
  document.body.append(dialog)
  const observer = observe()
  dialog.setAttribute('aria-hidden', 'true')
  dialog.style.pointerEvents = 'none'
  const records = observer.takeRecords()
  expect(dialog.isConnected).toBe(true)
  expect(records.some(r => qualifiesAck('dialog-close', r))).toBe(true)
  observer.disconnect()
})
it('does not acknowledge unrelated hiding or a dialog still accepting pointer input', () => {
  const dialog = document.createElement('div'), child = document.createElement('span')
  dialog.role = 'dialog'; dialog.append(child); document.body.append(dialog)
  const observer = observe()
  child.setAttribute('aria-hidden', 'true'); child.style.pointerEvents = 'none'
  expect(observer.takeRecords().some(r => qualifiesAck('dialog-close', r))).toBe(false)
  dialog.setAttribute('aria-hidden', 'true')
  expect(observer.takeRecords().some(r => qualifiesAck('dialog-close', r))).toBe(false)
  observer.disconnect()
})
it('still acknowledges immediate dialog removal without motion', () => {
  const dialog = document.createElement('div'); dialog.role = 'dialog'; document.body.append(dialog)
  const observer = observe(); dialog.remove()
  expect(observer.takeRecords().some(r => qualifiesAck('dialog-close', r))).toBe(true)
  observer.disconnect()
})
