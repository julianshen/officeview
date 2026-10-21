import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

const root = resolve('corpus') + '/'
const fetchSource = readFileSync(resolve('scripts/fetch-corpus.ts'), 'utf8')
const entries = [...fetchSource.matchAll(/lo\('([^']+)', '([a-f0-9]{40})', '([^']*)'\)/g)].map(m => ({ file: m[1], sha: m[2] }))
const expected = [
  {
    "file": "tdf104788.pptx",
    "sha": "319a6fb33c22156e83cebaf3447d7432beca8fb3",
    "blob": "c46fb3da50a96142d99b125370679482e87758d7",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "six slides and six notes bodies",
      "showNarration=1",
      "entrance presets 10/34/37/42/47/58",
      "animEffect, animScale and bldP"
    ]
  },
  {
    "file": "tdf115394.pptx",
    "sha": "a8a3e6a7fee5ac23bc2155b9391ead9402211147",
    "blob": "d01c90270f345e257c7b4d7bf0d5ad4685ed42cb",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "fade; speed-only and explicit durations 250/4250"
    ]
  },
  {
    "file": "tdf142915.pptx",
    "sha": "f8ddaaf0f5e1fb61e0d4404ea28757bc652ae4be",
    "blob": "889a08fd50a1f069410de5edb67b7b769e2ff283",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "advClick=0 advTm=1000; durations 0/10"
    ]
  },
  {
    "file": "tdf44223.pptx",
    "sha": "8a1321362a0229a25869e4e3d0422a5a51c5b5be",
    "blob": "6f0af688c2062607d261ecfcd84c39b347490b64",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "split with sound; paragraph background build; entrance preset 2"
    ]
  },
  {
    "file": "bnc591147.pptx",
    "sha": "a3ef3b658dea9b59ded319f2f092e73e4deaad24",
    "blob": "afa382e65770d9ac80e7b9a9c5b14cb727ff4d8e",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "interactive click sequence and media commands",
      "AVI and MP3 media with picture poster"
    ]
  },
  {
    "file": "connector-shape-animations.pptx",
    "sha": "8359e9974a5a43b094a7eecb73070b26fcb4df5a",
    "blob": "2fc2cee6af20ff9cee04e372103e6a2cc0e9ac90",
    "application": "Microsoft Macintosh PowerPoint",
    "pins": [
      "preset 22 targets connectors"
    ]
  },
  {
    "file": "tdf104792-smart-art-animation.pptx",
    "sha": "7c354d1bf1ca29a5ae9ee9d6f18dd68b90530043",
    "blob": "ac72639efd77ca2b64b9caa91052addb51f02b9e",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "SmartArt bldDgm"
    ]
  },
  {
    "file": "narration.pptx",
    "sha": "56f26399df894d05f24874a3cdc90694158e03bc",
    "blob": "5226ce6072158d6e738a05e571d6674f02742f47",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "audio, p14 media, M4A and transition"
    ]
  },
  {
    "file": "airplane-transition.pptx",
    "sha": "2c6eb9d95930a33318bcec785edf735aeb8adced",
    "blob": "18ee9aa1962107d954e8c87b51d6344961dd05c6",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "p15 airplane transition"
    ]
  },
  {
    "file": "tdf156808.pptx",
    "sha": "40641b868b20c73e694770c2c7ea71e232514fa1",
    "blob": "34256151b60bec1085e50ef5b5d5af7a823d8068",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "p14 flythrough"
    ]
  },
  {
    "file": "tdf149588_transparentSolidFill.pptx",
    "sha": "0cc9aaae5dc6655490513e8e4a93967f6248e23c",
    "blob": "8967590b92848a8d5f05f5fc599a17469e06d657",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "p159 morph"
    ]
  },
  {
    "file": "tdf105739.pptx",
    "sha": "319a6fb33c22156e83cebaf3447d7432beca8fb3",
    "blob": "08e26f55a91dda83422135719107a0a53535e597",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "slide gradient background"
    ]
  },
  {
    "file": "tdf146223.pptx",
    "sha": "4574a1ea408c2ac30042dca32d02207ec7add4da",
    "blob": "73d5eea144cce4c7e7e1a1a3a7d80d4de28b5a7d",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "slide and master picture backgrounds; master picture artwork"
    ]
  },
  {
    "file": "tdf123684.pptx",
    "sha": "943a534ac7cb3df513583e226c986dafd8ba246b",
    "blob": "87243422753bb3922828cbb382265d304cd34fb9",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "master bgRef 1003 and six non-placeholder artwork nodes"
    ]
  },
  {
    "file": "ooxtheme.pptx",
    "sha": "cd08835d1bd9274412c7326db32fb3752bc94fb3",
    "blob": "f0f3f2575894f715759a18439588cb149996140b",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "ten non-placeholder master artwork nodes"
    ]
  },
  {
    "file": "tdf93868.pptx",
    "sha": "f3d1ac75c4b7fa63022e54a9cbff46ba99535076",
    "blob": "f54fb0496e117eb55e0c16a023c045f80ba2b6dd",
    "application": "Microsoft Office PowerPoint",
    "pins": [
      "master gradient background and two non-placeholder shapes"
    ]
  }
] as Row[]
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
const P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main'
const P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
const rejected = ['effectOrder.pptx', 'animation-sound-embedded.pptx', 'animation-sound-external.pptx', 'tdf148478_transitionMusic.pptx']
type Row = { file: string; sha: string; blob: string; application: string; pins: string[] }
function manifest(): Row[] {
  expect(entries.map(r => r.file).sort()).toEqual(expected.map(r => r.file).sort())
  return entries.map(entry => {
    const pin = expected.find(row => row.file === entry.file)!
    expect(entry.sha, entry.file).toBe(pin.sha)
    return { ...pin, ...entry }
  })
}
it('every LibreOffice fetch entry has a pinned corpus lock hash', () => {
  const rows = manifest()
  expect(rows).toHaveLength(16)
  const lock = JSON.parse(readFileSync(root + 'corpus.lock.json', 'utf8'))
  for (const row of rows) expect(lock['lo-' + row.file], row.file).toMatch(/^[a-f0-9]{64}$/)
})
const absent = !expected.some(row => existsSync(root + 'lo-' + row.file))
const elements = (roots: (Document | Element)[], ns: string, name: string): Element[] =>
  roots.flatMap(r => Array.from(r.getElementsByTagNameNS(ns, name)))
const has = (roots: (Document | Element)[], ns: string, name: string) => expect(elements(roots, ns, name).length, name).toBeGreaterThan(0)
const attr = (roots: (Document | Element)[], ns: string, name: string, key: string, value: string) =>
  expect(elements(roots, ns, name).some(e => e.getAttribute(key) === value), `${name}@${key}=${value}`).toBe(true)
type Parts = { slides: Document[]; masters: Document[]; notes: Document[]; all: Document[]; zip: JSZip }
const masterArt = (p: Parts) => ['sp', 'pic', 'graphicFrame', 'cxnSp'].flatMap(n => elements(p.masters, P, n)).filter(e => !elements([e], P, 'ph').length)
const checks: Record<string, Record<string, (p: Parts) => void>> = {
  'tdf104788.pptx': {
    'six slides and six notes bodies': p => { expect(p.slides).toHaveLength(6); expect(p.notes).toHaveLength(6); for (const note of p.notes) expect(elements([note], P, 'sp').some(sp => elements([sp], P, 'ph').some(ph => ph.getAttribute('type') === 'body') && elements([sp], P, 'txBody').length > 0)).toBe(true) },
    'showNarration=1': p => attr(p.all, P, 'showPr', 'showNarration', '1'),
    'entrance presets 10/34/37/42/47/58': p => { for (const id of ['10','34','37','42','47','58']) expect(elements(p.slides,P,'cTn').some(e => e.getAttribute('presetClass') === 'entr' && e.getAttribute('presetID') === id)).toBe(true) },
    'animEffect, animScale and bldP': p => { for (const n of ['animEffect','animScale','bldP']) has(p.slides,P,n) },
  },
  'tdf115394.pptx': {
    'fade; speed-only and explicit durations 250/4250': p => { has(p.slides,P,'fade'); const t=elements(p.slides,P,'transition'); expect(t.some(e=>e.hasAttribute('spd') && !e.hasAttributeNS(P14,'dur'))).toBe(true); for(const v of ['250','4250']) expect(t.some(e=>e.getAttributeNS(P14,'dur')===v)).toBe(true) },
  },
  'tdf142915.pptx': {
    'advClick=0 advTm=1000; durations 0/10': p => { const t=elements(p.slides,P,'transition'); for(const v of ['0','10']) expect(t.some(e=>e.getAttribute('advClick')==='0' && e.getAttribute('advTm')==='1000' && e.getAttributeNS(P14,'dur')===v)).toBe(true) },
  },
  'tdf44223.pptx': {
    'split with sound; paragraph background build; entrance preset 2': p => { has(p.slides,P,'split'); has(p.slides,P,'sndAc'); attr(p.slides,P,'bldP','animBg','1'); expect(elements(p.slides,P,'cTn').some(e=>e.getAttribute('presetID')==='2' && e.getAttribute('presetClass')==='entr')).toBe(true) },
  },
  'bnc591147.pptx': {
    'interactive click sequence and media commands': p => { attr(p.slides,P,'cTn','nodeType','interactiveSeq'); attr(p.slides,P,'cond','evt','onClick'); attr(p.slides,P,'cTn','presetClass','mediacall'); has(p.slides,P,'cmd') },
    'AVI and MP3 media with picture poster': p => { expect(p.zip.file('ppt/media/video.avi')).not.toBeNull(); expect(p.zip.file('ppt/media/audio.mp3')).not.toBeNull(); has(p.slides,P,'pic'); has(p.slides,A,'videoFile') },
  },
  'connector-shape-animations.pptx': {
    'preset 22 targets connectors': p => { const ids=new Set(elements(elements(p.slides,P,'cxnSp'),P,'cNvPr').map(e=>e.getAttribute('id'))); const nodes=elements(p.slides,P,'cTn').filter(e=>e.getAttribute('presetID')==='22'); expect(nodes.length).toBeGreaterThan(0); expect(nodes.some(n => elements([n],P,'spTgt').some(e=>ids.has(e.getAttribute('spid'))))).toBe(true) },
  },
  'tdf104792-smart-art-animation.pptx': { 'SmartArt bldDgm': p => has(p.slides,A,'bldDgm') },
  'narration.pptx': { 'audio, p14 media, M4A and transition': p => { has(p.slides,P,'audio'); has(p.slides,P14,'media'); has(p.slides,P,'transition'); expect(p.zip.file('ppt/media/media1.m4a')).not.toBeNull() } },
  'airplane-transition.pptx': { 'p15 airplane transition': p => attr(p.slides,P15,'prstTrans','prst','airplane') },
  'tdf156808.pptx': { 'p14 flythrough': p => has(p.slides,P14,'flythrough') },
  'tdf149588_transparentSolidFill.pptx': { 'p159 morph': p => has(p.slides,P159,'morph') },
  'tdf105739.pptx': { 'slide gradient background': p => has(elements(p.slides,P,'bg'),A,'gradFill') },
  'tdf146223.pptx': { 'slide and master picture backgrounds; master picture artwork': p => { has(elements(p.slides,P,'bg'),A,'blipFill'); has(elements(p.masters,P,'bg'),A,'blipFill'); expect(masterArt(p).some(e=>e.localName==='pic')).toBe(true) } },
  'tdf123684.pptx': { 'master bgRef 1003 and six non-placeholder artwork nodes': p => { attr(p.masters,P,'bgRef','idx','1003'); expect(masterArt(p)).toHaveLength(6) } },
  'ooxtheme.pptx': { 'ten non-placeholder master artwork nodes': p => expect(masterArt(p)).toHaveLength(10) },
  'tdf93868.pptx': { 'master gradient background and two non-placeholder shapes': p => { has(elements(p.masters,P,'bg'),A,'gradFill'); expect(masterArt(p)).toHaveLength(2) } },
}
async function parts(zip: JSZip): Promise<Parts> {
  const all: Document[] = [], slides: Document[] = [], masters: Document[] = [], notes: Document[] = []
  for (const name of Object.keys(zip.files).filter(n => n.endsWith('.xml'))) {
    const doc = new DOMParser().parseFromString(await zip.file(name)!.async('string'), 'application/xml')
    expect(doc.getElementsByTagName('parsererror')).toHaveLength(0)
    all.push(doc)
    if (/^ppt\/slides\/slide\d+\.xml$/.test(name)) slides.push(doc)
    if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name)) masters.push(doc)
    if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)) notes.push(doc)
  }
  return { zip, all, slides, masters, notes }
}

describe.skipIf(absent)(`slideshow fixture provenance${absent ? ' — run bun scripts/fetch-corpus.ts' : ''}`, () => {
  it('every manifest file exists, matches its blob sha and reports a Microsoft producer', async () => {
    const rows = manifest()
    expect(rows.map(r=>r.file).sort()).toEqual(Object.keys(checks).sort())
    for(const row of rows) {
      expect(row.sha).toMatch(/^[a-f0-9]{40}$/)
      expect(row.blob).toMatch(/^[a-f0-9]{40}$/)
      const bytes=readFileSync(root+'lo-'+row.file)
      const blob=createHash('sha1').update(`blob ${bytes.length}\0`).update(bytes).digest('hex')
      expect(blob,row.file).toBe(row.blob)
      const zip=await JSZip.loadAsync(bytes)
      const app=new DOMParser().parseFromString(await zip.file('docProps/app.xml')!.async('string'),'application/xml')
      const producer=app.getElementsByTagNameNS('*','Application')[0]?.textContent
      expect(producer,row.file).toBe(row.application)
      expect(producer).toContain('Microsoft')
    }
  })
  it('each fixture contains the elements its pins claim', async () => {
    for(const row of manifest()) {
      expect(row.pins,row.file).toEqual(Object.keys(checks[row.file]))
      const p=await parts(await JSZip.loadAsync(readFileSync(root+'lo-'+row.file)))
      for(const check of Object.values(checks[row.file])) check(p)
    }
  })
  it('PROVENANCE lists every manifest file and every rejected candidate', () => {
    const rows=manifest()
    expect(existsSync(resolve('docs/validation/slideshow-corpus.md'))).toBe(true)
    const doc=readFileSync(resolve('docs/validation/slideshow-corpus.md'),'utf8')
    for(const row of rows) { expect(doc).toContain(row.file); expect(doc).toContain(row.sha); expect(doc).toContain(row.blob) }
    for(const file of rejected) { expect(doc).toContain(file); expect(existsSync(root+'lo-'+file)).toBe(false) }
    expect(doc).toContain('MPL-2.0')
    expect(doc).toContain('metadata-reported')
  })
})
