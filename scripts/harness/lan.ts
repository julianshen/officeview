import { networkInterfaces, type NetworkInterfaceInfo } from 'node:os'

/**
 * First external IPv4 address. When none exists the result is flagged
 * `loopbackFallback` so the CLI can say loudly that no other device can reach
 * the printed URL.
 */
export function detectLanIp(interfaces: NodeJS.Dict<NetworkInterfaceInfo[]> = networkInterfaces()): { ip: string; loopbackFallback: boolean } {
  for (const list of Object.values(interfaces)) {
    for (const address of list ?? []) {
      if ((address.family === 'IPv4' || (address.family as unknown) === 4) && !address.internal) {
        return { ip: address.address, loopbackFallback: false }
      }
    }
  }
  return { ip: '127.0.0.1', loopbackFallback: true }
}
