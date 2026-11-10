import { createHmac, randomInt } from 'node:crypto'
import { codeword40 } from './bch'
/** Prototype wire choice: unsigned uint32 BE keyGen followed by the five
 * MSB-first payload bytes (token24, salt8, CRC8). Output is 40 binary bits. */
export function messageFor(key: Uint8Array, gen: number, token: number, salt: number): Uint8Array {
  if (key.length !== 32 || !Number.isSafeInteger(gen) || gen < 0 || gen > 0xffffffff) throw new Error('Invalid message key/generation')
  const bits = codeword40(token, salt), input = Buffer.alloc(9)
  input.writeUInt32BE(gen)
  for (let i=0;i<40;i++) input[4+(i>>3)]! |= bits[i]! << (7-i%8)
  const digest = createHmac('sha256',key).update(input).digest()
  return Uint8Array.from({length:40},(_,i)=>(digest[i>>3]!>>(7-i%8))&1)
}
/** The caller owns the audit set for this key generation. Bounded retries
 * refuse rather than publishing a colliding message (including a stuck RNG). */
export function insertMessage(key: Uint8Array, gen: number, token: number, used: Set<string>, draw:()=>number = ()=>randomInt(256)) {
  for(let attempt=0;attempt<256;attempt++) {
    const salt8=draw(), msg40=messageFor(key,gen,token,salt8), id=Array.from(msg40).join('')
    if(used.has(id)) continue
    used.add(id); return {salt8,msg40}
  }
  throw new Error('Message salt retries exhausted')
}
