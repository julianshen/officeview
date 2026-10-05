import { describe, expect, it } from 'vitest'
import { mkdtempSync, existsSync, writeFileSync, rmSync, mkdirSync, copyFileSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'
import { parseXmlOrdered } from '../src/core/xml'
import { parseGeometry, resolveGeometry, resolvePreset, type GeometryDefinition, type ResolvedCommand } from '../src/drawing/geometry'
import presets from '../src/drawing/presets.json'

const geometry = (xml: string) => parseGeometry(parseXmlOrdered(`<spPr xmlns:a="urn:a">${xml}</spPr>`))
const pathGeometry = (commands: string, guides = '', pathAttrs = '') => geometry(`<a:custGeom><a:gdLst>${guides}</a:gdLst><a:pathLst><a:path ${pathAttrs}>${commands}</a:path></a:pathLst></a:custGeom>`)
const point = (command: ResolvedCommand): number[] => command.slice(1).filter((v): v is number => typeof v === 'number')

it('resolves a rectangle in local pixels and retains its text rectangle', () => {
  const result = resolvePreset('rect', 200, 100)
  expect(result.paths[0].commands).toEqual([['moveTo', 0, 0], ['lnTo', 200, 0], ['lnTo', 200, 100], ['lnTo', 0, 100], ['close']])
  expect(result.textRect).toEqual({ left: 0, top: 0, right: 200, bottom: 100 })
  expect(result.issues).toEqual([])
})

it('keeps the twelve dodecagon vertices in source order', () => {
  const commands = resolvePreset('dodecagon', 216, 216).paths[0].commands
  expect(commands.map(point)).toEqual([[0, 79.06], [28.94, 28.94], [79.06, 0], [136.94, 0], [187.06, 28.94], [216, 79.06], [216, 136.94], [187.06, 187.06], [136.94, 216], [79.06, 216], [28.94, 187.06], [0, 136.94], []])
})

it('retains preset adjustments and resolves negative connector adjustments', () => {
  const definition = geometry('<a:prstGeom prst="bentConnector4"><a:avLst><a:gd name="adj1" fmla="val -10000"/><a:gd name="adj2" fmla="val 25000"/></a:avLst></a:prstGeom>')
  expect(definition.preset).toBe('bentConnector4')
  expect(definition.adjustments).toEqual([['adj1', 'val -10000'], ['adj2', 'val 25000']])
  expect(resolveGeometry(definition, 200, 100).paths[0].commands).toEqual([['moveTo', 0, 0], ['lnTo', -20, 0], ['lnTo', -20, 25], ['lnTo', 200, 25], ['lnTo', 200, 100]])
  expect(resolvePreset('bentConnector4', 200, 100).paths[0].commands).toEqual([['moveTo', 0, 0], ['lnTo', 100, 0], ['lnTo', 100, 50], ['lnTo', 200, 50], ['lnTo', 200, 100]])
})

it('uses ellipse polar angles for the chord start and end', () => {
  const result = resolvePreset('chord', 200, 100)
  const [move, arc] = result.paths[0].commands
  expect(move[0]).toBe('moveTo')
  expect(move[1]).toBeCloseTo(100 + 100 / Math.sqrt(5))
  expect(move[2]).toBeCloseTo(50 + 100 / Math.sqrt(5))
  expect(arc[0]).toBe('arcTo')
  if (arc[0] !== 'arcTo') throw new Error('Expected arc')
  const [, cx, cy, rx, ry, start, end, anticlockwise] = arc
  expect([cx, cy, rx, ry]).toEqual([100, 50, 100, 50])
  expect(start).toBeCloseTo(Math.atan2(2, 1))
  expect(cx + rx * Math.cos(end)).toBeCloseTo(100)
  expect(cy + ry * Math.sin(end)).toBeCloseTo(0)
  expect(anticlockwise).toBe(false)
})

it('resolves the smiley mouth quadratic control from the pinned definition', () => {
  const mouth = resolvePreset('smileyFace', 216, 100).paths[2]
  expect(mouth.fill).toBe('none')
  expect(mouth.commands[0][0]).toBe('moveTo')
  expect(mouth.commands[0][1]).toBeCloseTo(216 * 4969 / 21699)
  expect(mouth.commands[0][2]).toBeCloseTo(16515 / 216 - 4.653)
  expect(mouth.commands[1][0]).toBe('quadBezTo')
  expect(point(mouth.commands[1])).toEqual([108, 16515 / 216 + 3 * 4.653, 166.4, 16515 / 216 - 4.653])
})

it('keeps custom command order, coordinate scaling, flags, and multiple subpaths', () => {
  const definition = pathGeometry('<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo><a:quadBezTo><a:pt x="20" y="0"/><a:pt x="30" y="10"/></a:quadBezTo><a:lnTo><a:pt x="40" y="0"/></a:lnTo><a:close/><a:moveTo><a:pt x="50" y="10"/></a:moveTo><a:cubicBezTo><a:pt x="60" y="0"/><a:pt x="70" y="10"/><a:pt x="80" y="0"/></a:cubicBezTo>', '', 'w="100" h="20" fill="darkenLess" stroke="false"')
  const path = resolveGeometry(definition, 200, 100).paths[0]
  expect(path.fill).toBe('darkenLess')
  expect(path.stroke).toBe(false)
  expect(path.commands).toEqual([['moveTo', 0, 0], ['lnTo', 20, 50], ['quadBezTo', 40, 0, 60, 50], ['lnTo', 80, 0], ['close'], ['moveTo', 100, 50], ['cubicBezTo', 120, 0, 140, 50, 160, 0]])
})

it('defaults absent geometry to rect and preserves an explicit unknown preset', () => {
  expect(parseGeometry(parseXmlOrdered('<spPr/>')).preset).toBe('rect')
  const unknown = geometry('<a:prstGeom prst="unrecognizedShape"/>')
  expect(unknown.preset).toBe('unrecognizedShape')
  const result = resolveGeometry(unknown, 20, 10)
  expect(result.paths).toEqual([])
  expect(result.issues).toEqual([expect.objectContaining({ kind: 'unknown-preset' })])
})

it('skips a malformed path while retaining valid neighbors and audit issues', () => {
  const definition = geometry('<a:custGeom><a:gdLst><a:gd name="bad" fmla="*/ w 1 0"/></a:gdLst><a:pathLst><a:path><a:moveTo><a:pt x="missing" y="0"/></a:moveTo></a:path><a:path><a:moveTo><a:pt x="bad" y="0"/></a:moveTo></a:path><a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="w" y="h"/></a:lnTo></a:path></a:pathLst></a:custGeom>')
  const result = resolveGeometry(definition, 40, 20)
  expect(result.paths).toHaveLength(1)
  expect(result.paths[0].commands).toEqual([['moveTo', 0, 0], ['lnTo', 40, 20]])
  expect(result.issues.filter(issue => issue.kind === 'invalid-path').map(issue => issue.pathIndex)).toEqual([0, 1])
  expect(result.issues.some(issue => issue.kind === 'invalid-guide')).toBe(true)
})

it('keeps the sweep direction and full circle through fixed path scaling', () => {
  const definition = pathGeometry('<a:moveTo><a:pt x="100" y="50"/></a:moveTo><a:arcTo wR="50" hR="25" stAng="0" swAng="-21600000"/><a:lnTo><a:pt x="0" y="50"/></a:lnTo>', '', 'w="100" h="100"')
  const arc = resolveGeometry(definition, 200, 100).paths[0].commands[1]
  expect(arc).toEqual(['arcTo', 100, 50, 100, 25, 0, -2 * Math.PI, true])
})

const guideCases: Array<[string, number]> = [
  ['val 12', 12], ['+- 7 5 2', 10], ['*/ 7 6 3', 14], ['+/ 7 5 3', 4], ['?: 1 8 9', 8], ['?: 0 8 9', 9], ['abs -6', 6], ['at2 0 3', 5400000], ['cat2 10 3 4', 6], ['cos 10 cd4', 0], ['max 3 4', 4], ['min 3 4', 3], ['mod 2 3 6', 7], ['pin 2 9 5', 5], ['sat2 10 3 4', 8], ['sin 10 cd4', 10], ['sqrt 49', 7], ['tan 10 cd8', 10], ['val 3cd4', 16200000], ['val wd32', 10], ['val ssd8', 20], ['val ls', 320], ['val hc', 160], ['val vc', 80],
]
describe('guide formulas', () => {
  it.each(guideCases)('%s', (formula, expected) => {
    const definition = pathGeometry('<a:moveTo><a:pt x="result" y="0"/></a:moveTo>', `<a:gd name="result" fmla="${formula.replace(/&/g, '&amp;')}"/>`)
    expect(resolveGeometry(definition, 320, 160).paths[0].commands[0][1]).toBeCloseTo(expected)
  })
  it.each(['wat 1', 'val unrecognized', '*/ 1 1 0', 'sqrt -1', 'val Infinity', 'val 1x', 'val 1 2'])('rejects invalid formula %s', formula => {
    const definition = pathGeometry('<a:moveTo><a:pt x="result" y="0"/></a:moveTo>', `<a:gd name="result" fmla="${formula}"/>`)
    const result = resolveGeometry(definition, 100, 50)
    expect(result.paths).toEqual([])
    expect(result.issues.some(issue => issue.kind === 'invalid-guide')).toBe(true)
  })
})

it('resolves all 187 pinned presets to finite nonempty paths at square, wide, and tall extents', () => {
  expect(Object.keys(presets)).toHaveLength(187)
  for (const name of Object.keys(presets)) {
    for (const [width, height] of [[100, 100], [320, 80], [80, 320]]) {
      const result = resolvePreset(name, width, height)
      expect(result.paths.length, `${name} ${width}x${height}: ${JSON.stringify(result.issues)}`).toBeGreaterThan(0)
      expect(result.issues, `${name} ${width}x${height}`).toEqual([])
      for (const path of result.paths) {
        expect(path.commands.length).toBeGreaterThan(0)
        for (const command of path.commands) expect(point(command).every(Number.isFinite)).toBe(true)
      }
    }
  }
})

it('handles zero and invalid extents without nonfinite path coordinates', () => {
  const definition = parseGeometry(undefined)
  for (const [width, height] of [[0, 0], [100, 0], [NaN, 100], [-1, 100], [Infinity, 100]]) {
    const result = resolveGeometry(definition, width, height)
    expect(result.paths.flatMap(path => path.commands).every(command => point(command).every(Number.isFinite))).toBe(true)
  }
})

// This compile-time assignment also pins the normalized source-data contract.
const normalized: GeometryDefinition = { adjustments: [], guides: [], paths: [{ commands: [['moveTo', '0', '0']] }] }
void normalized


it('normalizes only the redundant zero operands in the pinned POI circular-arrow guides', () => {
  for (const preset of ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow']) {
    const guides = (presets as unknown as Record<string, GeometryDefinition>)[preset].guides
    expect(guides.find(([name]) => name === 'xB')?.[1]).toBe('+- xH 0 dxB')
    expect(guides.find(([name]) => name === 'yB')?.[1]).toBe('+- yH 0 dyB')
  }
  const guides = (presets as unknown as Record<string, GeometryDefinition>).leftRightCircularArrow.guides
  expect(guides.find(([name]) => name === 'xJ')?.[1]).toBe('+- xI 0 dxJ')
  expect(guides.find(([name]) => name === 'yJ')?.[1]).toBe('+- yI 0 dyJ')
})

it('rejects an unpinned local preset source without creating an output file', () => {
  const dir = mkdtempSync(join(tmpdir(), 'officeview-preset-guard-'))
  try {
    const input = join(dir, 'wrong.xml')
    const output = join(dir, 'presets.json')
    writeFileSync(input, '<presetShapeDefinitons/>')
    const result = spawnSync('bun', ['scripts/generate-drawing-presets.ts', input, output], { cwd: process.cwd(), encoding: 'utf8' })
    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('Pinned preset input SHA-256 mismatch')
    expect(existsSync(output)).toBe(false)
  } finally { rmSync(dir, { recursive: true, force: true }) }
})


it('can validate generator input before a generated preset catalog exists', () => {
  const dir = mkdtempSync(join(tmpdir(), 'officeview-preset-bootstrap-'))
  try {
    mkdirSync(join(dir, 'scripts'))
    mkdirSync(join(dir, 'src/core'), { recursive: true })
    mkdirSync(join(dir, 'src/drawing'))
    copyFileSync('scripts/generate-drawing-presets.ts', join(dir, 'scripts/generate-drawing-presets.ts'))
    copyFileSync('src/core/xml.ts', join(dir, 'src/core/xml.ts'))
    copyFileSync('src/drawing/geometry-xml.ts', join(dir, 'src/drawing/geometry-xml.ts'))
    symlinkSync(join(process.cwd(), 'node_modules'), join(dir, 'node_modules'))
    const input = join(dir, 'wrong.xml')
    writeFileSync(input, '<presetShapeDefinitons/>')
    const result = spawnSync('bun', ['scripts/generate-drawing-presets.ts', input], { cwd: dir, encoding: 'utf8' })
    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('Pinned preset input SHA-256 mismatch')
  } finally { rmSync(dir, { recursive: true, force: true }) }
})

it('applies adjustment overrides before evaluating the ordered guide list', () => {
  const definition = geometry('<a:custGeom><a:avLst><a:gd name="adj" fmla="val 10"/></a:avLst><a:gdLst><a:gd name="adj" fmla="+- adj 1 0"/><a:gd name="x" fmla="*/ adj 2 1"/></a:gdLst><a:pathLst><a:path><a:moveTo><a:pt x="x" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>')
  expect(resolveGeometry(definition, 100, 100, { adj: 20 }).paths[0].commands).toEqual([['moveTo', 42, 0]])
})

it.each([
  ['pin 10 20 5', 5],
  ['pin 10 10 5', 5],
  ['pin 10 7 5', 10],
])('follows ordered pin branches with inverted bounds: %s', (formula, expected) => {
  const definition = pathGeometry('<a:moveTo><a:pt x="result" y="0"/></a:moveTo>', `<a:gd name="result" fmla="${formula}"/>`)
  expect(resolveGeometry(definition, 100, 100).paths[0].commands).toEqual([['moveTo', expected, 0]])
})

it('evaluates guides and built-ins separately in each fixed path coordinate space', () => {
  const definition: GeometryDefinition = {
    adjustments: [],
    guides: [['quarter', '*/ h 1 4'], ['shortHalf', '*/ ss 1 2']],
    paths: [
      { width: '100', height: '20', commands: [['moveTo', 'hc', 'quarter'], ['lnTo', 'w', 'h']] },
      { width: '20', height: '100', commands: [['moveTo', 'shortHalf', 'quarter'], ['lnTo', 'w', 'h']] },
      { commands: [['moveTo', 'hc', 'quarter'], ['lnTo', 'w', 'h']] },
    ],
    textRect: ['hc', 'quarter', 'w', 'h'],
  }
  const result = resolveGeometry(definition, 200, 100)
  expect(result.paths.map(path => path.commands)).toEqual([
    [['moveTo', 100, 25], ['lnTo', 200, 100]],
    [['moveTo', 100, 25], ['lnTo', 200, 100]],
    [['moveTo', 100, 25], ['lnTo', 200, 100]],
  ])
  expect(result.textRect).toEqual({ left: 100, top: 25, right: 200, bottom: 100 })
  expect(result.issues).toEqual([])
})

it('audits a guide failure in one path coordinate space while preserving another space', () => {
  const definition: GeometryDefinition = {
    adjustments: [],
    guides: [['local', '*/ 1 1 h']],
    paths: [
      { width: '100', commands: [['moveTo', 'local', '0']] },
      { width: '100', height: '20', commands: [['moveTo', 'local', '0'], ['lnTo', 'w', 'h']] },
    ],
  }
  const result = resolveGeometry(definition, 200, 0)
  expect(result.paths.map(path => path.commands)).toEqual([[['moveTo', 0.1, 0], ['lnTo', 200, 0]]])
  expect(result.issues).toContainEqual(expect.objectContaining({ kind: 'invalid-guide', guide: 'local', pathIndex: 0 }))
  expect(result.issues).toContainEqual(expect.objectContaining({ kind: 'invalid-path', pathIndex: 0 }))
  expect(result.issues.some(issue => issue.pathIndex === 1)).toBe(false)
})

it('evaluates replacement adjustment formulas in their document order', () => {
  const definition = geometry('<a:prstGeom prst="bentConnector4"><a:avLst><a:gd name="adj2" fmla="val 25000"/><a:gd name="adj1" fmla="val adj2"/></a:avLst></a:prstGeom>')
  expect(definition.adjustments).toEqual([['adj2', 'val 25000'], ['adj1', 'val adj2']])
  const result = resolveGeometry(definition, 200, 100)
  expect(result.paths[0]?.commands).toEqual([['moveTo', 0, 0], ['lnTo', 50, 0], ['lnTo', 50, 25], ['lnTo', 200, 25], ['lnTo', 200, 100]])
  expect(result.issues).toEqual([])
})

it('keeps non-overridden adjustment defaults before ordered document overrides', () => {
  const definition = geometry('<a:prstGeom prst="bentConnector4"><a:avLst><a:gd name="adj2" fmla="val adj1"/></a:avLst></a:prstGeom>')
  expect(definition.adjustments).toEqual([['adj1', 'val 50000'], ['adj2', 'val adj1']])
  const result = resolveGeometry(definition, 200, 100)
  expect(result.paths[0]?.commands).toEqual([['moveTo', 0, 0], ['lnTo', 100, 0], ['lnTo', 100, 50], ['lnTo', 200, 50], ['lnTo', 200, 100]])
  expect(result.issues).toEqual([])
})

it('accepts an empty custom path list without a geometry issue', () => {
  const result = resolveGeometry(geometry('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:pathLst/></a:custGeom>'), 200, 100)
  expect(result.paths).toEqual([])
  expect(result.issues).toEqual([])
})
