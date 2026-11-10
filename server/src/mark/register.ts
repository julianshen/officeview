import { createCanvas } from '@napi-rs/canvas'
import { checkBitmap, type Bitmap, type Chip } from './types'
export const lum = (b: Bitmap, x: number, y: number): number => {
  const i = (y * b.width + x) * 4, a = b.data[i + 3]! / 255
  return a * (0.2126 * b.data[i]! + 0.7152 * b.data[i + 1]! + 0.0722 * b.data[i + 2]!) + (1 - a) * 255
}
interface Offset { x: number; y: number; error: number }
/** Smoothed bitmap resample (Skia high quality), the same operation the Stage-0 `scale` attack applies. */
function resize(b: Bitmap, factor: number): Bitmap {
  checkBitmap(b)
  if (!Number.isFinite(factor) || factor <= 0) throw new Error('Invalid scale')
  const w = Math.round(b.width * factor), h = Math.round(b.height * factor)
  if (!Number.isSafeInteger(w) || !Number.isSafeInteger(h) || w < 1 || h < 1) throw new Error('Invalid scaled size')
  const src = createCanvas(b.width, b.height), sctx = src.getContext('2d'), pixels = sctx.createImageData(b.width, b.height)
  pixels.data.set(b.data); sctx.putImageData(pixels, 0, 0)
  const c = createCanvas(w, h), ctx = c.getContext('2d')
  ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high'; ctx.drawImage(src, 0, 0, w, h)
  return { width: w, height: h, data: new Uint8ClampedArray(ctx.getImageData(0, 0, w, h).data) }
}
/** Exhaustive coarse level and top-64 refinement at every level; block
 * averages avoid point-sampling aliasing. Full-resolution certification below
 * rejects a pose only by a nonnegative squared-error lower bound. */
function peaks(ref: Bitmap, cap: Bitmap): { x: number; y: number; stride: number }[] {
 let stride=1;while(Math.max(ref.width,ref.height)/stride>256)stride*=2
 let best:Offset[]=[]
 for(;stride>=1;stride/=2){
  const r=plane(ref,stride),c=plane(cap,stride),seen=new Set<number>(),next:Offset[]=[]
  const maxX=ref.width-cap.width,maxY=ref.height-cap.height
  const visit=(x:number,y:number)=>{
   x=Math.max(0,Math.min(maxX,x));y=Math.max(0,Math.min(maxY,y))
   const id=y*ref.width+x;if(seen.has(id))return;seen.add(id)
   let error=0,n=0;const jump=Math.max(1,Math.floor(Math.sqrt(c.w*c.h/4096)))
   for(let cy=0;cy<c.h-1;cy+=jump)for(let cx=0;cx<c.w-1;cx+=jump){
    const rx=Math.min(r.w-1,Math.round(x/stride)+cx),ry=Math.min(r.h-1,Math.round(y/stride)+cy)
    error+=(c.data[cy*c.w+cx]!-r.data[ry*r.w+rx]!)**2;n++
   }
   next.push({x,y,error:error/Math.max(1,n)})
  }
  if(!best.length){for(let y=0;y<=maxY+stride;y+=stride)for(let x=0;x<=maxX+stride;x+=stride)visit(x,y)}
  else for(const p of best)for(let dy=-2;dy<=2;dy++)for(let dx=-2;dx<=2;dx++)visit(p.x+dx*stride,p.y+dy*stride)
  next.sort((a,b)=>a.error-b.error);best=next.slice(0,64)
 }
 return best.map(p=>({...p,stride:1}))
}
function plane(b:Bitmap,stride:number){
 const w=Math.ceil(b.width/stride),h=Math.ceil(b.height/stride),data=new Float64Array(w*h)
 for(let by=0;by<h;by++)for(let bx=0;bx<w;bx++){
  let sum=0,n=0;for(let y=by*stride;y<Math.min(b.height,(by+1)*stride);y++)for(let x=bx*stride;x<Math.min(b.width,(bx+1)*stride);x++){sum+=lum(b,x,y);n++}
  data[by*w+bx]=sum/n
 }
 return {w,h,data}
}
/** Certify the translation minimum. Sparse pages use the exact SSD identity
 * C² + R² - 2CR about their background, avoiding billions of white pixels.
 * Rich pages reject by nonnegative partial squared-error sums. */
function exactMinimum(ref:Bitmap,cap:Bitmap,seed:Offset):Offset {
 const r=plane(ref,1),c=plane(cap,1),n=c.data.length
 let lo=Infinity,hi=-Infinity;for(const v of r.data){lo=Math.min(lo,v);hi=Math.max(hi,v)}
 if(hi-lo<1e-6)return seed
 const background=255,sparse:{x:number;y:number;v:number}[]=[]
 for(let i=0;i<r.data.length;i++)if(Math.abs(r.data[i]!-background)>1e-6){sparse.push({x:i%r.w,y:Math.floor(i/r.w),v:r.data[i]!-background});if(sparse.length>4096)break}
 if(sparse.length<=4096){
  let baseError=0;for(const v of c.data)baseError+=(v-background)**2
  let best=seed,limit=seed.error*n
  for(let y=0;y<=ref.height-cap.height;y++)for(let x=0;x<=ref.width-cap.width;x++){
   let sum=baseError
   for(const p of sparse)if(p.x>=x&&p.y>=y&&p.x<x+c.w&&p.y<y+c.h)sum+=p.v*p.v-2*p.v*(c.data[(p.y-y)*c.w+p.x-x]!-background)
   if(sum<limit-1e-7){limit=sum;best={x,y,error:sum/n}}
  }
  return best
 }
 // Integral reference energy gives a Cauchy-Schwarz SSD lower bound.
 // Blank windows are exact in O(1); do not rescan millions of white pixels.
 const iw=r.w+1,energy=new Float64Array(iw*(r.h+1))
 for(let y=0;y<r.h;y++){let row=0;for(let x=0;x<r.w;x++){row+=(r.data[y*r.w+x]!-background)**2;energy[(y+1)*iw+x+1]=energy[y*iw+x+1]!+row}}
 let captureEnergy=0;for(const v of c.data)captureEnergy+=(v-background)**2
 const edge=(i:number)=>Math.abs(c.data[i]!-c.data[Math.min(n-1,i+1)]!)+Math.abs(c.data[i]!-c.data[Math.min(n-1,i+c.w)]!)
 const probes:number[]=[],step=Math.max(1,Math.floor(n/4096))
 for(let i=0;i<n;i+=step){let at=i;for(let j=i+1;j<Math.min(n,i+step);j++)if(edge(j)>edge(at))at=j;probes.push(at)}
 probes.sort((a,b)=>edge(b)-edge(a))
 const offsets=probes.map(i=>Math.floor(i/c.w)*r.w+i%c.w),marked=new Uint8Array(n)
 for(const i of probes)marked[i]=1
 let best=seed,limit=seed.error*n
 for(let y=0;y<=ref.height-cap.height;y++)for(let x=0;x<=ref.width-cap.width;x++){
  const referenceEnergy=Math.max(0,energy[(y+c.h)*iw+x+c.w]!-energy[y*iw+x+c.w]!-energy[(y+c.h)*iw+x]!+energy[y*iw+x]!)
  if((Math.sqrt(referenceEnergy)-Math.sqrt(captureEnergy))**2>limit+1e-7)continue
  if(referenceEnergy<1e-8){if(captureEnergy<limit-1e-7){limit=captureEnergy;best={x,y,error:limit/n}}continue}
  let sum=0;const base=y*r.w+x
  for(let j=0;j<probes.length;j++){sum+=(c.data[probes[j]!]!-r.data[base+offsets[j]!]!)**2;if(sum>limit)break}
  if(sum>limit)continue
  outer:for(let cy=0;cy<c.h;cy++)for(let cx=0;cx<c.w;cx++){
   const i=cy*c.w+cx;if(marked[i])continue
   sum+=(c.data[i]!-r.data[base+cy*r.w+cx]!)**2;if(sum>limit)break outer
  }
  if(sum<limit){limit=sum;best={x,y,error:sum/n}}
 }
 return best
}
/** Ambiguous means separated poses (>=2 pixels) within max(.1, 5% of best)
 * luminance MSE. Only keyed support is used, NEVER candidate symbols or Z.
 * Rank residual energy per supported pixel; nulls use this exact procedure. */
function supportSeeds(ref:Bitmap,cap:Bitmap,poses:Offset[],chips:Chip[],sx:number,sy:number):Offset[]{
 const first=poses[0];if(!first)return []
 if(!poses.some(p=>Math.hypot(p.x-first.x,p.y-first.y)>=2&&p.error-first.error<=Math.max(.1,first.error*.05)))return []
 const width=cap.width+1,integral=new Float64Array(width*(cap.height+1))
 for(let y=0;y<cap.height;y++){let row=0;for(let x=0;x<cap.width;x++){row+=(lum(cap,x,y)-lum(ref,x+first.x,y+first.y))**2;integral[(y+1)*width+x+1]=integral[y*width+x+1]!+row}}
 const energy=(x:number,y:number)=>{let sum=0,area=0
  for(const chip of chips){const x0=Math.ceil(chip.x*sx-x),y0=Math.ceil(chip.y*sy-y),x1=Math.floor((chip.x+chip.cell)*sx-x),y1=Math.floor((chip.y+chip.cell)*sy-y)
   if(x0<0||y0<0||x1>cap.width||y1>cap.height||x1<=x0||y1<=y0)continue
   sum+=integral[y1*width+x1]!-integral[y0*width+x1]!-integral[y1*width+x0]!+integral[y0*width+x0]!;area+=(x1-x0)*(y1-y0)
  }return area?sum/area:0}
 const step=Math.max(1,Math.floor((chips[0]?.cell??4)*Math.min(sx,sy)/2)),best:Offset[]=[]
 for(let y=0;y<=ref.height-cap.height;y+=step)for(let x=0;x<=ref.width-cap.width;x+=step)best.push({x,y,error:-energy(x,y)})
 best.sort((a,b)=>a.error-b.error)
 const refined:Offset[]=[]
 for(const p of best.slice(0,8))for(let dy=-step;dy<=step;dy++)for(let dx=-step;dx<=step;dx++){
  const x=p.x+dx,y=p.y+dy;if(x<0||y<0||x>ref.width-cap.width||y>ref.height-cap.height)continue
  refined.push({x,y,error:-energy(x,y)})
 }
 refined.sort((a,b)=>a.error-b.error);return refined.slice(0,8)
}
/** Where a capture sits on its reference: resample by `factor`, crop at
 * (x, y) in ORIGINAL reference pixels (scaled-first crops may be fractional),
 * `scaleX/Y` being capture pixels per original reference pixel. */
export type Pose = { factor: number; x: number; y: number; scaleX: number; scaleY: number }
export interface Registration extends Pose {
  /** Already aligned/resampled to capture dimensions, never sampled out of bounds. */
  reference: Bitmap; error: number
}
function rectangle(b: Bitmap, x: number, y: number, width: number, height: number): Bitmap {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let row = 0; row < height; row++) data.set(b.data.subarray(((y + row) * b.width + x) * 4, ((y + row) * b.width + x + width) * 4), row * width * 4)
  return { width, height, data }
}
function sample(b: Bitmap, x: number, y: number): number {
  x = Math.max(0, Math.min(b.width - 1, x)); y = Math.max(0, Math.min(b.height - 1, y))
  const x0 = Math.floor(x), y0 = Math.floor(y), tx = x - x0, ty = y - y0
  return (1 - ty) * ((1 - tx) * lum(b, x0, y0) + tx * lum(b, Math.min(x0 + 1, b.width - 1), y0)) + ty * ((1 - tx) * lum(b, x0, Math.min(y0 + 1, b.height - 1)) + tx * lum(b, Math.min(x0 + 1, b.width - 1), Math.min(y0 + 1, b.height - 1)))
}
/** The reference as `register` aligns it to a width x height capture at
 * `pose`. Two constructions exist and are told apart by the pose itself:
 * scale-then-crop (scale is exactly round(extent*factor)/extent and the crop
 * origin is a whole scaled pixel) or crop-then-scale (integer origin, scale
 * is capture/extent). Poses satisfying both read as scale-then-crop. */
export function alignReference(reference: Bitmap, pose: Pose, width: number, height: number): Bitmap {
  checkBitmap(reference)
  const { factor, x, y, scaleX, scaleY } = pose
  if (![factor, x, y, scaleX, scaleY].every(Number.isFinite) || factor <= 0 || scaleX <= 0 || scaleY <= 0 || !Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) throw new RangeError('Invalid pose')
  const w = Math.round(reference.width * factor), h = Math.round(reference.height * factor)
  const sx = x * scaleX, sy = y * scaleY
  const scaledFirst = scaleX === w / reference.width && scaleY === h / reference.height && Math.abs(sx - Math.round(sx)) < 1e-6 && Math.abs(sy - Math.round(sy)) < 1e-6
  if (scaledFirst) {
    const ox = Math.round(sx), oy = Math.round(sy)
    if (ox < 0 || oy < 0 || ox + width > w || oy + height > h) throw new RangeError('Pose outside the reference')
    return rectangle(factor === 1 ? reference : resize(reference, factor), ox, oy, width, height)
  }
  const cw = Math.round(width / scaleX), ch = Math.round(height / scaleY)
  if (!Number.isInteger(x) || !Number.isInteger(y) || x < 0 || y < 0 || cw < 1 || ch < 1 || x + cw > reference.width || y + ch > reference.height) throw new RangeError('Pose outside the reference')
  return resize(rectangle(reference, x, y, cw, ch), factor)
}
export function register(capture: Bitmap, reference: Bitmap, chips: Chip[] = []): Registration | undefined {
  let best: Registration | undefined
  const step = 1
  const errorOf = (read: (x: number, y: number) => number, stride = step): number => {
    let sum = 0, count = 0
    for (let y = 0; y < capture.height; y += stride) for (let x = 0; x < capture.width; x += stride) { sum += (lum(capture, x, y) - read(x, y)) ** 2; count++ }
    return sum / count
  }
  for (const factor of [1, 0.5, 2]) {
    const w = Math.round(reference.width * factor), h = Math.round(reference.height * factor), slack = Math.ceil(factor)
    if (capture.width > w || capture.height > h || capture.width < Math.floor(w * 0.75) - slack || capture.height < Math.floor(h * 0.75) - slack) continue
    const ref = factor === 1 ? reference : resize(reference, factor)
    const seeds = w === capture.width && h === capture.height ? [{ x: 0, y: 0, stride: 1 }] : [...peaks(ref, capture), {x:Math.floor((w-capture.width)/2),y:Math.floor((h-capture.height)/2),stride:1}]
    // Preserve scale-then-crop support, where origins are scaled pixel integers.
    let scaledBest: { x: number; y: number; error: number } | undefined
    const visited = new Set<number>()
    for (const seed of seeds) for (let y = Math.max(0, seed.y - seed.stride); y <= Math.min(h - capture.height, seed.y + seed.stride); y++) for (let x = Math.max(0, seed.x - seed.stride); x <= Math.min(w - capture.width, seed.x + seed.stride); x++) {
      const id = y * w + x; if (visited.has(id)) continue; visited.add(id)
      const error = errorOf((cx, cy) => lum(ref, cx + x, cy + y), Math.max(1,Math.floor(Math.sqrt(capture.width*capture.height/4096))))
      if (!scaledBest || error < scaledBest.error) scaledBest = { x, y, error }
    }
    if(scaledBest){
      scaledBest.error=errorOf((cx,cy)=>lum(ref,cx+scaledBest!.x,cy+scaledBest!.y))
      // Cheap ranking is only an upper-bound seed, never the final objective.
      if(seeds.length>1){
        const centre={x:Math.floor((w-capture.width)/2),y:Math.floor((h-capture.height)/2),error:0}
        centre.error=errorOf((cx,cy)=>lum(ref,cx+centre.x,cy+centre.y))
        if(centre.error<scaledBest.error)scaledBest=centre
        scaledBest=exactMinimum(ref,capture,scaledBest)
      }
    }
    if(chips.length && seeds.length>1){
      const ranked=[...(scaledBest?[scaledBest]:[]),...seeds.slice(0,8),seeds[seeds.length-1]!].map(p=>({...p,error:errorOf((cx,cy)=>lum(ref,cx+p.x,cy+p.y))})).sort((a,b)=>a.error-b.error)
      const synced=supportSeeds(ref,capture,ranked,chips,w/reference.width,h/reference.height)
      if(synced.length){
        const candidate=synced[0]!,error=errorOf((cx,cy)=>lum(ref,cx+candidate.x,cy+candidate.y))
        if(error<=ranked[0]!.error+Math.max(.1,ranked[0]!.error*.05)){
          scaledBest={x:candidate.x,y:candidate.y,error}
          seeds.unshift(...synced.map(p=>({...p,stride:1})))
        }
      }
    }
    if (scaledBest && (!best || scaledBest.error < best.error)) {
      const { x, y, error } = scaledBest
      best = { factor, x: x * reference.width / w, y: y * reference.height / h, scaleX: w / reference.width, scaleY: h / reference.height, reference: rectangle(ref, x, y, capture.width, capture.height), error }
    }
    if (factor === 1) continue // both orders coincide
    // Invert round(sourceExtent*factor): half-scale has two possible integer
    // extents. Search integer ORIGINAL offsets, including half-pixel phases
    // after scaling. Coarse interpolation ranks candidates; actual napi
    // crop-then-resize determines the final content error and reference.
    const extents = (size: number, max: number) => {
      const result: number[] = []
      for (let n = Math.ceil((size - 0.5) / factor); n < (size + 0.5) / factor; n++) if (n > 0 && n <= max) result.push(n)
      return result
    }
    for (const cw of extents(capture.width, reference.width)) for (const ch of extents(capture.height, reference.height)) {
      const sx = capture.width / cw, sy = capture.height / ch
      const shortlist: { x: number; y: number; error: number }[] = [], seen = new Set<number>()
      for (const seed of seeds) {
        const radius = Math.ceil(seed.stride / factor) + 1, ox = Math.round(seed.x / factor), oy = Math.round(seed.y / factor)
        for (let y = Math.max(0, oy - radius); y <= Math.min(reference.height - ch, oy + radius); y++) for (let x = Math.max(0, ox - radius); x <= Math.min(reference.width - cw, ox + radius); x++) {
          const id = y * reference.width + x; if (seen.has(id)) continue; seen.add(id)
          const error = errorOf((cx, cy) => sample(reference, x + (cx + 0.5) / sx - 0.5, y + (cy + 0.5) / sy - 0.5), Math.max(step, Math.floor(Math.sqrt(capture.width * capture.height / 256))))
          if (shortlist.length < 8 || error < shortlist[shortlist.length - 1]!.error) { shortlist.push({ x, y, error }); shortlist.sort((a, b) => a.error - b.error); if (shortlist.length > 8) shortlist.pop() }
        }
      }
      for (const { x, y } of shortlist) {
        const aligned = resize(rectangle(reference, x, y, cw, ch), factor)
        const error = errorOf((cx, cy) => lum(aligned, cx, cy))
        if (!best || error < best.error) best = { factor, x, y, scaleX: sx, scaleY: sy, reference: aligned, error }
      }
    }
  }
  return best
}
