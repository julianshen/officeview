#!/usr/bin/env python3
"""Generate the pinned OFFICIAL TEXT-warp catalog TS module.

Tracked reproducible generator (stdlib only, no dependencies).

Usage (from the repository root):
  python3 scripts/warp/generate-text-warp-catalog.py \
    [--src tests/fixtures/presetTextWarpDefinitions.xml] \
    [--out src/drawing/text-warp-catalog.ts] [--check]

Input (READ-ONLY, hash-pinned): tests/fixtures/presetTextWarpDefinitions.xml
  (ECMA-376 Part 1 5th edition DrawingML preset text warp definitions,
  SHA-256 89640bd4aeedfa2c578d1d1acc50d569865ebdef315fd4b468087c25e36acada).
Output: src/drawing/text-warp-catalog.ts (verified byte-identical by --check).

Tokens preserved verbatim (guide references like 'x1', arc attrs 'wd2'/'adval');
avLst official defaults kept as FORMULAS ('val cd2'); ahLst bounds kept as raw
tokens for the resolver to evaluate; pathLst commands keep source order.
Text-warp geometry seeds its own guide set (l,t,r,b,w,h,hc,vc,wd2,hd2,wd3,ss,cd2)
per ECMA-376 presetTextWarpDefinitions - NOT the shape catalog.
"""
import re, hashlib, json, sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
SRC = os.path.join(ROOT, 'tests', 'fixtures', 'presetTextWarpDefinitions.xml')
OUT = os.path.join(ROOT, 'src', 'drawing', 'text-warp-catalog.ts')
for _index, _arg in enumerate(sys.argv):
    if _arg == '--src' and _index + 1 < len(sys.argv):
        SRC = sys.argv[_index + 1]
    if _arg == '--out' and _index + 1 < len(sys.argv):
        OUT = sys.argv[_index + 1]
SUPPORTED = ['textArchUp','textArchDown','textCircle','textWave1','textWave2',
             'textInflate','textDeflate','textSlantUp','textSlantDown','textCurveUp','textCurveDown']

raw = open(SRC, 'rb').read()
sha = hashlib.sha256(raw).hexdigest()
PIN = '89640bd4aeedfa2c578d1d1acc50d569865ebdef315fd4b468087c25e36acada'
if sha != PIN:
    print('FATAL: pinned source hash mismatch', sha)
    sys.exit(1)
xml = raw.decode('utf8')


def block_of(name):
    m = re.search(r'<%s>(.*?)</%s>' % (name, name), xml, re.S)
    if not m:
        raise SystemExit('preset missing: ' + name)
    return m.group(1)


def named_guides(body, tag):
    if '<' + tag not in body:
        return []
    seg = body.split('<' + tag, 1)[1].split('</' + tag, 1)[0]
    return [(m.group(1), m.group(2))
            for m in re.finditer(r'<gd name="(\w+)" fmla="([^"]+)"\s*/?>', seg)]


def handles_of(body):
    out = []
    for hm in re.finditer(r'<(ahPolar|ahXY)\b([^>]*?)(?:/>|>(?:(?:(?!ahPolar|ahXY).)*?)</\1>)', body, re.S):
        a = dict(re.findall(r'(gdRef\w+|min\w+|max\w+)="([^"]*)"', hm.group(2)))
        guide = a.get('gdRefAng') or a.get('gdRefR') or a.get('gdRefY') or a.get('gdRefX')
        mnkey = [k for k in a if k.startswith('min')][0]
        mxkey = [k for k in a if k.startswith('max')][0]
        axis = {'minAng': 'angle', 'minR': 'radius', 'minY': 'y', 'minX': 'x'}[mnkey]
        out.append((guide, axis, a[mnkey], a[mxkey]))
    return out


def paths_of(body):
    pl = re.search(r'<pathLst[^>]*>(.*)</pathLst>', body, re.S)
    if not pl:
        return []
    out = []
    for pm in re.finditer(r'<path([^>]*?)/?>(.*?)</path>', pl.group(1), re.S):
        attrs = dict(re.findall(r'\b(w|h)="([^"]*)"', pm.group(1)))
        commands = []
        for cm in re.finditer(r'<(moveTo|lnTo|quadBezTo|cubicBezTo|arcTo|close)([^>]*?)(?:/>|>(.*?)</\1>)',
                              pm.group(2), re.S):
            kind, carg, inner = cm.group(1), cm.group(2), cm.group(3) or ''
            if kind == 'arcTo':
                a = dict(re.findall(r'(wR|hR|stAng|swAng)="([^"]*)"', carg))
                commands.append(('arcTo', None, (a['wR'], a['hR'], a['stAng'], a['swAng'])))
            elif kind == 'close':
                commands.append(('close', (), None))
            else:
                pts = re.findall(r'<pt x="([^"]+)" y="([^"]+)"\s*/?>', inner)
                commands.append((kind, tuple(v for pt in pts for v in pt), None))
        out.append((attrs, commands))
    return out


def tok(v):
    s = str(v)
    if re.fullmatch(r'-?\d+', s):
        return s
    if re.fullmatch(r'-?\d+\.\d+', s):
        return s
    return json.dumps(s)


lines = []
lines.append('/**')
lines.append(' * Pinned OFFICIAL TEXT-warp geometry catalog (ECMA-376 Part 1, 5th edition, December 2016).')
lines.append(' *')
lines.append(' * Source: ECMA-376-1_5th_edition_december_2016.zip > OfficeOpenXML-DrawingMLGeometries.zip')
lines.append(' *         > presetTextWarpDefinitions.xml')
lines.append(' * Source SHA-256 (bytes): %s' % sha)
lines.append(' * Reference: https://ecma-international.org/publications-and-standards/standards/ecma-376/')
lines.append(' *')
lines.append(' * TEXT-warp catalog (a:bodyPr > a:prstTxWarp). Deliberately NOT routed through the shape')
lines.append(' * preset catalog (presets.json): the text-warp path context seeds its own guide set')
lines.append(' * (l,t,r,b,w,h,hc,vc,wd2,hd2,wd3,ss,cd2) per ECMA-376; reusing shape validation would')
lines.append(' * evaluate the wrong seed set.')
lines.append(' *')
lines.append(' * Tokens are preserved verbatim; avLst keeps official DEFAULT FORMULAS (e.g.')
lines.append(' * textArchUp adj = "val cd2"); gdLst keeps derived guide order; ahLst provides the')
lines.append(' * official handle bounds; pathLst keeps command order with guide-reference tokens.')
lines.append(' *')
lines.append(' * Regenerated by scripts/warp/generate-text-warp-catalog.py; regeneration asserts the pinned')
lines.append(' * hash. Generated content is a normalized transcription of the pinned official XML')
lines.append(' * (attribution: ECMA-376).')
lines.append(' */')
lines.append('')
lines.append('export interface TextWarpCommand {')
lines.append("  kind: 'moveTo' | 'lnTo' | 'quadBezTo' | 'cubicBezTo' | 'arcTo' | 'close'")
lines.append('  /** Point tokens (guide references or numeric literals), xy interleaved. */')
lines.append('  values?: ReadonlyArray<string | number>')
lines.append('  /** arcTo carries its four tokens as attributes (DrawingML angle units). */')
lines.append('  attrs?: Partial<Record<\'wR\' | \'hR\' | \'stAng\' | \'swAng\', string | number>>')
lines.append('}')
lines.append('')
lines.append('export interface TextWarpDefinition {')
lines.append('  /** avLst: official adjustment DEFAULTS as guide formulas. */')
lines.append('  adjustments: ReadonlyArray<readonly [string, string]>')
lines.append('  /** gdLst: ordered derived guides. */')
lines.append('  guides: ReadonlyArray<readonly [string, string]>')
lines.append('  /** ahLst: official handle (adjustment) bounds. */')
lines.append("  handles: ReadonlyArray<{ guide: string; axis: 'angle' | 'radius' | 'y' | 'x'; min: string | number; max: string | number }>")
lines.append('  /** pathLst: authoritative boundary commands. */')
lines.append("  paths: ReadonlyArray<{ attrs: Partial<Record<'w' | 'h', string | number>>; commands: ReadonlyArray<TextWarpCommand> }>")
lines.append('}')
lines.append('')
lines.append('export const OFFICIAL_TEXT_WARP_REFERENCE = {')
lines.append("  file: 'presetTextWarpDefinitions.xml',")
lines.append("  sha256: '" + sha + "',")
lines.append("  source: 'ECMA-376 Part 1 5th edition (December 2016), DrawingML preset text warp definitions',")
lines.append('} as const')
lines.append('')
lines.append('export const TEXT_WARP_CATALOG: Readonly<Record<string, TextWarpDefinition>> = {')
for name in SUPPORTED:
    body = block_of(name)
    avl = named_guides(body, 'avLst') or named_guides(body, 'gdLst')
    gdl = named_guides(body, 'gdLst')
    handles = handles_of(body)
    paths = paths_of(body)
    lines.append('  ' + json.dumps(name) + ': {')
    lines.append('    adjustments: [')
    for nm, fm in avl:
        lines.append('      [%s, %s],' % (json.dumps(nm), json.dumps(fm)))
    lines.append('    ],')
    lines.append('    guides: [')
    for nm, fm in gdl:
        lines.append('      [%s, %s],' % (json.dumps(nm), json.dumps(fm)))
    lines.append('    ],')
    lines.append('    handles: [')
    for guide, axis, mn, mx in handles:
        lines.append('      { guide: %s, axis: %s, min: %s, max: %s },'
                     % (json.dumps(guide), json.dumps(axis), tok(mn), tok(mx)))
    lines.append('    ],')
    lines.append('    paths: [')
    for attrs, commands in paths:
        pa = ', '.join('%s: %s' % (k, tok(v)) for k, v in attrs.items())
        lines.append('      { attrs: { %s }, commands: [' % pa)
        for kind, values, arcattrs in commands:
            if kind == 'arcTo':
                wR, hR, stAng, swAng = arcattrs
                lines.append("        { kind: 'arcTo', attrs: { wR: %s, hR: %s, stAng: %s, swAng: %s } },"
                             % (tok(wR), tok(hR), tok(stAng), tok(swAng)))
            elif kind == 'close':
                lines.append("        { kind: 'close' },")
            else:
                vals = ', '.join(tok(v) for v in values)
                lines.append('        { kind: %s, values: [%s] },' % (json.dumps(kind), vals))
        lines.append('      ] },')
    lines.append('    ],')
    lines.append('  },')
lines.append('}')
lines.append('')

content = '\n'.join(lines)
if '--check' in sys.argv:
    current = open(OUT, 'r', encoding='utf8').read()
    if current != content:
        print('MISMATCH: %s differs from generator output' % OUT)
        sys.exit(1)
    print('OK: %s matches generator output (%d bytes)' % (OUT, len(content)))
else:
    open(OUT, 'w').write(content)
    print('wrote %s (%d bytes); presets: %s' % (OUT, len(content), ', '.join(SUPPORTED)))
