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
 // Block lower bounds, coarse to fine: over any block, the squared-error sum
 // is at least (capture sum - reference sum)^2 / area (Cauchy-Schwarz), so
 // any set of blocks bounds a window's SSD from below. Reference block sums
 // come from an integral image at every offset. A nearly flat page (uniform
 // fill, faint tiled mark) leaves the energy bound and the edge probes
 // nothing to reject by; the blocks still separate wrong offsets cheaply.
 // Blocks deviating most from the capture mean go first.
 const sums=new Float64Array(iw*(r.h+1))
 for(let y=0;y<r.h;y++){let row=0;for(let x=0;x<r.w;x++){row+=r.data[y*r.w+x]!-128;sums[(y+1)*iw+x+1]=sums[y*iw+x+1]!+row}}
 let captureMean=0;for(const v of c.data)captureMean+=v;captureMean/=n
 const levels=[16,8,4].filter(s=>s*2<=Math.min(c.w,c.h)).map(s=>{
  const bw=Math.floor(c.w/s),bh=Math.floor(c.h/s),order:number[]=[],total=new Float64Array(bw*bh)
  for(let by=0;by<bh;by++)for(let bx=0;bx<bw;bx++){let t=0;for(let y=by*s;y<(by+1)*s;y++)for(let x=bx*s;x<(bx+1)*s;x++)t+=c.data[y*c.w+x]!-128;total[by*bw+bx]=t;order.push(by*bw+bx)}
  order.sort((a,b)=>Math.abs(total[b]!-(captureMean-128)*s*s)-Math.abs(total[a]!-(captureMean-128)*s*s))
  const at=new Int32Array(order.map(i=>Math.floor(i/bw)*s*iw+i%bw*s)),cs=new Float64Array(order.map(i=>total[i]!))
  return {s,at,cs,corner:s*iw+s,down:s*iw}
 })
 const blocked=(x:number,y:number,limit:number):boolean=>{
  const base=y*iw+x,stop=limit*(1+1e-9)+1e-6
  for(const {s,at,cs,corner,down} of levels){
   let sum=0;const area=s*s
   for(let j=0;j<at.length;j++){const k=base+at[j]!,d=cs[j]!-(sums[k+corner]!-sums[k+s]!-sums[k+down]!+sums[k]!);sum+=d*d/area;if(sum>stop)return true}
  }
  return false
 }
 let best=seed,limit=seed.error*n
 for(let y=0;y<=ref.height-cap.height;y++)for(let x=0;x<=ref.width-cap.width;x++){
  const referenceEnergy=Math.max(0,energy[(y+c.h)*iw+x+c.w]!-energy[y*iw+x+c.w]!-energy[(y+c.h)*iw+x]!+energy[y*iw+x]!)
  if((Math.sqrt(referenceEnergy)-Math.sqrt(captureEnergy))**2>limit+1e-7)continue
  if(referenceEnergy<1e-8){if(captureEnergy<limit-1e-7){limit=captureEnergy;best={x,y,error:limit/n}}continue}
  if(blocked(x,y,limit))continue
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
export type Pose = { factor: number; x: number; y: number; scaleX: number; scaleY: number; order: 'scale-crop' | 'crop-scale' }
export interface Registration extends Pose {
  /** Already aligned/resampled to capture dimensions, never sampled out of bounds. */
  reference: Bitmap
  /** Luminance MSE of the capture against `reference`. */
  error: number
  /** `error` over (luminance variance of `reference` + RESIDUAL_FLOOR); above
   * GEOMETRY_ERROR_MAX the capture's geometry is out of scope. */
  residual: number
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
 * `pose`, by the construction `pose.order` names: scale-crop (resample the
 * whole reference by `factor`, then crop at the whole scaled pixel x*scaleX,
 * y*scaleY) or crop-scale (crop at the integer original origin, extent
 * capture/scale, then resample by `factor`). */
export function alignReference(reference: Bitmap, pose: Pose, width: number, height: number): Bitmap {
  checkBitmap(reference)
  const { factor, x, y, scaleX, scaleY } = pose
  if (![factor, x, y, scaleX, scaleY].every(Number.isFinite) || factor <= 0 || scaleX <= 0 || scaleY <= 0 || !Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) throw new RangeError('Invalid pose')
  const w = Math.round(reference.width * factor), h = Math.round(reference.height * factor)
  if (pose.order !== 'scale-crop' && pose.order !== 'crop-scale') throw new RangeError('Invalid pose')
  if (pose.order === 'scale-crop') {
    const sx = x * scaleX, sy = y * scaleY
    if (scaleX !== w / reference.width || scaleY !== h / reference.height || Math.abs(sx - Math.round(sx)) >= 1e-6 || Math.abs(sy - Math.round(sy)) >= 1e-6) throw new RangeError('Invalid scale-crop pose')
    const ox = Math.round(sx), oy = Math.round(sy)
    if (ox < 0 || oy < 0 || ox + width > w || oy + height > h) throw new RangeError('Pose outside the reference')
    return rectangle(factor === 1 ? reference : resize(reference, factor), ox, oy, width, height)
  }
  const cw = Math.round(width / scaleX), ch = Math.round(height / scaleY)
  if (!Number.isInteger(x) || !Number.isInteger(y) || x < 0 || y < 0 || cw < 1 || ch < 1 || x + cw > reference.width || y + ch > reference.height) throw new RangeError('Pose outside the reference')
  return resize(rectangle(reference, x, y, cw, ch), factor)
}
/** Browser zoom range the factor search covers (capture px per reference px). */
export const ZOOM_MIN = 0.45, ZOOM_MAX = 2.2
/** Normalised residual (see `residual`) above which a capture's geometry is
 * out of scope: rotation, perspective, a photograph. 1.5 x the largest
 * in-scope residual of server/test/mark-geometry.test.ts, measured
 * 2026-11-10: 0.06152, a marked blank 256x256 page (its own mark energy;
 * pypptx-test.pptx unit 0 peaks at 0.004, q70; the Stage-0 synthetic text
 * page at q70 is 0.0224). Out of scope on pypptx-test unit 0: rotate 2 deg
 * 0.843, rotate 5 deg 1.146, keystone 0.05 0.842. Recalibrate with the
 * detector calibration task. */
export const GEOMETRY_ERROR_MAX = 0.0923
/** Registration MSE over the aligned reference's luminance variance plus a
 * floor, so a sparse page is not judged on its mark energy alone. */
function residual(aligned: Bitmap, error: number): number {
  let sum = 0, sq = 0
  const n = aligned.width * aligned.height
  for (let y = 0; y < aligned.height; y++) for (let x = 0; x < aligned.width; x++) { const v = lum(aligned, x, y); sum += v; sq += v * v }
  const variance = Math.max(0, sq / n - (sum / n) ** 2)
  return error / (variance + RESIDUAL_FLOOR)
}
/** (amplitude + pedestal)^2 at DEFAULT_PARAMS: the largest squared luminance
 * change the mark itself makes, so a blank page's mark energy stays in scope. */
const RESIDUAL_FLOOR = 16
/** Seeds above this residual skip the exhaustive translation certificate. */
const CERTIFY_MAX = 4 * GEOMETRY_ERROR_MAX
/** The factor an uncropped resample implies: round(extent*f) reproduces both
 * capture extents. Midpoint of the admissible interval, if any. */
function impliedFactor(capture: Bitmap, reference: Bitmap): number | undefined {
  const lo = Math.max((capture.width - 0.5) / reference.width, (capture.height - 0.5) / reference.height)
  const hi = Math.min((capture.width + 0.5) / reference.width, (capture.height + 0.5) / reference.height)
  if (!(lo < hi)) return undefined
  const f = (lo + hi) / 2
  return f >= ZOOM_MIN && f <= ZOOM_MAX && Math.round(reference.width * f) === capture.width && Math.round(reference.height * f) === capture.height ? f : undefined
}
/** Coarse-to-fine scale search for a cropped zoom. Level 0 is an exhaustive
 * multiplicative factor grid x exhaustive translation on ~160 px thumbnails
 * (both sides through the same resampler); each finer level doubles the
 * resolution, halves the factor step and searches +-4 steps / +-4 px around
 * the best three distinct hypotheses. Deterministic; returns the
 * full-resolution best factor, which `register` then evaluates in both orders. */
function zoomSearch(capture: Bitmap, reference: Bitmap): number | undefined {
  const W = reference.width, H = reference.height, cw = capture.width, ch = capture.height
  const lo = Math.max(ZOOM_MIN, cw / W, ch / H), hi = Math.min(ZOOM_MAX, (cw + 1) / (0.75 * W), (ch + 1) / (0.75 * H))
  if (!(lo <= hi)) return undefined
  type Cand = { f: number; x: number; y: number; error: number }
  const score = (cap: { w: number; h: number; data: Float64Array }, f: number, c: number, xs: [number, number], ys: [number, number]): Cand | undefined => {
    const r = plane(resize(reference, f * c), 1)
    const maxX = r.w - cap.w, maxY = r.h - cap.h
    if (maxX < -1 || maxY < -1) return undefined
    const jump = Math.max(1, Math.floor(Math.sqrt(cap.w * cap.h / 4096)))
    let best: Cand | undefined
    for (let y = Math.max(0, ys[0]); y <= Math.min(Math.max(0, maxY), ys[1]); y++) for (let x = Math.max(0, xs[0]); x <= Math.min(Math.max(0, maxX), xs[1]); x++) {
      let error = 0, n = 0
      for (let cy = 0; cy < cap.h; cy += jump) for (let cx = 0; cx < cap.w; cx += jump) {
        error += (cap.data[cy * cap.w + cx]! - r.data[Math.min(r.h - 1, y + cy) * r.w + Math.min(r.w - 1, x + cx)]!) ** 2; n++
      }
      error /= n
      if (!best || error < best.error) best = { f, x, y, error }
    }
    return best
  }
  let c = Math.min(1, 160 / Math.max(cw, ch))
  const thumb = (k: number) => plane(k === 1 ? capture : resize(capture, k), 1)
  let cap = thumb(c), step = 0.5 / Math.max(cap.w, cap.h), level: Cand[] = []
  for (let f = lo; f <= hi * (1 + step); f *= 1 + step) { const s = score(cap, Math.min(f, hi), c, [0, Infinity], [0, Infinity]); if (s) level.push(s) }
  // Distinct (resampled extent, offset) candidates only: neighbouring factors
  // that round to the same thumbnail are one hypothesis, not three.
  const best3 = (found: Cand[], c: number): Cand[] => {
    const seen = new Set<string>(), out: Cand[] = []
    for (const s of [...found].sort((a, b) => a.error - b.error)) {
      const id = `${Math.round(W * s.f * c)}x${Math.round(H * s.f * c)}@${s.x},${s.y}`
      if (!seen.has(id)) { seen.add(id); out.push(s) }
      if (out.length === 3) break
    }
    return out
  }
  // Seed with the best local minima of the coarse factor profile.
  let top = best3(level.filter((s, i) => (i === 0 || s.error <= level[i - 1]!.error) && (i === level.length - 1 || s.error <= level[i + 1]!.error)), c)
  while (c < 1 && top.length) {
    const next = Math.min(1, c * 2), ratio = next / c
    c = next; cap = thumb(c); step = 0.5 / Math.max(cap.w, cap.h)
    const found: Cand[] = []
    for (const t of top) for (let j = -4; j <= 4; j++) {
      const f = Math.min(hi, Math.max(lo, t.f * (1 + step) ** j)), x = Math.round(t.x * ratio * f / t.f), y = Math.round(t.y * ratio * f / t.f)
      const s = score(cap, f, c, [x - 4, x + 4], [y - 4, y + 4]); if (s) found.push(s)
    }
    top = best3(found, c)
  }
  return top[0]?.f
}
/** Residual energy per keyed-support pixel through a pose: the measure
 * `supportSeeds` ranks ambiguous poses by, mapped as `scoreAligned` maps chips. */
function supportEnergy(capture: Bitmap, pose: Pose, aligned: Bitmap, chips: Chip[]): number {
  let sum = 0, area = 0
  for (const chip of chips) {
    const x0 = Math.ceil((chip.x - pose.x) * pose.scaleX), y0 = Math.ceil((chip.y - pose.y) * pose.scaleY)
    const x1 = Math.floor((chip.x + chip.cell - pose.x) * pose.scaleX), y1 = Math.floor((chip.y + chip.cell - pose.y) * pose.scaleY)
    if (x0 < 0 || y0 < 0 || x1 > capture.width || y1 > capture.height || x1 <= x0 || y1 <= y0) continue
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) sum += (lum(capture, x, y) - lum(aligned, x, y)) ** 2
    area += (x1 - x0) * (y1 - y0)
  }
  return area ? sum / area : 0
}
/** `supportSeeds`' ambiguity band: within max(.1, 5 % of the best) luminance MSE. */
const ambiguous = (error: number, best: number): boolean => error - best <= Math.max(.1, best * .05)
export function register(capture: Bitmap, reference: Bitmap, chips: Chip[] = []): Registration | undefined {
  checkBitmap(capture); checkBitmap(reference)
  type Hypothesis = Omit<Registration, 'residual'> & { exact: boolean }
  const hypotheses: Hypothesis[] = []
  const step = 1
  const errorOf = (read: (x: number, y: number) => number, stride = step): number => {
    let sum = 0, count = 0
    for (let y = 0; y < capture.height; y += stride) for (let x = 0; x < capture.width; x += stride) { sum += (lum(capture, x, y) - read(x, y)) ** 2; count++ }
    return sum / count
  }
  const tried = new Set<number>()
  /** One factor's best pose, as a hypothesis competing with the others.
   * `exact`: a Stage-0 factor, which wins otherwise-equal ties. */
  const atFactor = (factor: number, exact = false): void => {
    if (tried.has(factor)) return
    tried.add(factor)
    // Certify only a seed that can win outright, by more than the ambiguity
    // band, over every hypothesis so far: a rival inside the band is decided
    // by keyed support, not by its exact MSE minimum, and certifying it is
    // the hours (a nearly flat page's SSD surface is flat to JPEG noise).
    const before = Math.min(...hypotheses.map(h => h.error))
    let best: Omit<Registration, 'residual'> | undefined
    const w = Math.round(reference.width * factor), h = Math.round(reference.height * factor), slack = Math.ceil(factor)
    if (capture.width > w || capture.height > h || capture.width < Math.floor(w * 0.75) - slack || capture.height < Math.floor(h * 0.75) - slack) return
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
        // Certify only a plausible factor: an exhaustive minimum under a
        // far-out-of-scope bound barely prunes (minutes on a wrong zoom).
        if(!ambiguous(before,scaledBest.error)&&residual(rectangle(ref,scaledBest.x,scaledBest.y,capture.width,capture.height),scaledBest.error)<=CERTIFY_MAX)scaledBest=exactMinimum(ref,capture,scaledBest)
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
    if (scaledBest) {
      const { x, y, error } = scaledBest
      best = { factor, x: x * reference.width / w, y: y * reference.height / h, scaleX: w / reference.width, scaleY: h / reference.height, order: 'scale-crop', reference: rectangle(ref, x, y, capture.width, capture.height), error }
    }
    const done = () => { if (best) hypotheses.push({ ...best, exact }) }
    if (factor === 1) return done() // both orders coincide
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
        if (!best || error < best.error) best = { factor, x, y, scaleX: sx, scaleY: sy, order: 'crop-scale', reference: aligned, error }
      }
    }
    done()
  }
  // Hypotheses inside the ambiguity band of the best fit are decided by keyed
  // support, exactly as `supportSeeds` decides poses within one factor; else
  // the lowest MSE. Exact factors, then evaluation order, break equal ties.
  const result = (): Registration | undefined => {
    if (!hypotheses.length) return undefined
    const least = Math.min(...hypotheses.map(h => h.error))
    const band = hypotheses.filter(h => ambiguous(h.error, least))
    const rank = band.length > 1 && chips.length ? band.map(h => -supportEnergy(capture, h, h.reference, chips)) : band.map(h => h.error)
    let at = 0
    for (let i = 1; i < band.length; i++) if (rank[i]! < rank[at]! || (rank[i] === rank[at] && band[i]!.exact && !band[at]!.exact)) at = i
    const { exact: _, ...best } = band[at]!
    return { ...best, residual: residual(best.reference, best.error) }
  }
  const inScope = (): Registration | undefined => { const r = result(); return r && r.residual <= GEOMETRY_ERROR_MAX ? r : undefined }
  // Every admissible hypothesis competes; none stops the search early. An
  // uncropped zoom of 0.75 (1.5) has exactly the extents of a 25 % crop at
  // factor 1 (2), and on a nearly flat page the wrong one is in scope too.
  // The uncropped resample the extents imply goes first: it is a single pose,
  // so it needs no translation search, and an exact factor whose seed sits in
  // its ambiguity band is not certified. Only when nothing fits in scope does
  // the continuous zoom search run.
  const exactFactors = [1, 0.5, 2], implied = impliedFactor(capture, reference)
  if (implied !== undefined && !exactFactors.some(f => Math.round(reference.width * f) === capture.width && Math.round(reference.height * f) === capture.height)) atFactor(implied)
  for (const factor of exactFactors) atFactor(factor, true)
  const found = inScope()
  if (found) return found
  const zoom = zoomSearch(capture, reference)
  if (zoom !== undefined) atFactor(zoom)
  return result()
}
