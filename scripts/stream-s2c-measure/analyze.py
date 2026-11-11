import json,statistics as st,sys
V=sys.argv[1] if len(sys.argv)>1 else 'A'
d=json.load(open('s2c-out.json'));R=d['rows']
def pct(a,b): return 100*(1-a/b) if b else 0
print('rows',len(R),'docs',len(set(r['file'] for r in R)))
diffAB=sum(1 for r in R if r['culled']['A'].get('lean')!=r['culled']['B'].get('lean'));print('A!=B rows',diffAB)
for r in R:
  if not r['culled']['A']['qualifies']: print('nq',r['file'],r['unit'],r['scale'],r.get('finalUnmarked'),r['culled']['A'].get('reason'),r['culled']['A'].get('unmarkedBlocks'),r['culled']['A'].get('blocks'))
def summ(rows,key,label,clip=True):
  F=sum(r['finalRaster'] for r in rows)
  C=[]
  for r in rows:
    c=r['culled'][V]; q=c['qualifies'] and c.get(key) is not None
    v=c[key] if q else r['finalRaster']
    if clip: v=min(v,r['finalRaster'])
    C.append(v)
  per=[pct(c,r['finalRaster']) for c,r in zip(C,rows)]
  return F,sum(C),pct(sum(C),F),st.median(per) if per else 0
groups={'all':R}
for f in ['pptx','docx','xlsx']: groups[f]=[r for r in R if r['format']==f]
groups['pptx@1']=[r for r in R if r['format']=='pptx' and r['scale']==1]
groups['pptx@2']=[r for r in R if r['format']=='pptx' and r['scale']==2]
for g,rows in groups.items():
  q=[r for r in rows if r['culled']['A']['qualifies']]
  print(f"\n== {g}: units {len(rows)} qualifying {len(q)} ({100*len(q)/len(rows):.1f}%)")
  for key in ['total','lean']:
    for clip in [True,False]:
      F,C,p,m=summ(q,key,'',clip); Fa,Ca,pa,ma=summ(rows,key,'',clip)
      print(f"  {key:5} clip={clip!s:5} qual: final {F} culled {C} save {p:.1f}% median {m:.1f}% | all: save {pa:.1f}% median {ma:.1f}%")
  F=sum(r['finalRaster'] for r in rows); Rl=sum(min(r['relaxed']['bytes'],r['finalRaster']) for r in rows)
  per=[pct(min(r['relaxed']['bytes'],r['finalRaster']),r['finalRaster']) for r in rows]
  nr=sum(1 for r in rows if r['relaxed']['kind']=='commands')
  print(f"  relaxed commands: {nr}/{len(rows)} ship as commands; total save {pct(Rl,F):.1f}% median {st.median(per):.1f}%")
  NV=sum(r['finalNoVisibleMark'] for r in rows); print(f"  final-raster w/o visible mark: save {pct(NV,F):.1f}% (median {st.median([pct(r['finalNoVisibleMark'],r['finalRaster']) for r in rows]):.1f}%)")
  wins=sum(1 for r in q if r['culled']['A']['lean']<r['finalRaster']); wins_s=sum(1 for r in q if r['culled']['A']['total']<r['finalRaster'])
  print(f"  qualifying units where culled < final: lean {wins}/{len(q)}, strict {wins_s}/{len(q)}")
  print(f"  totals: final {F}, overlay-raster sum {sum(r['culled']['A'].get('overlay') or 0 for r in q)}, rasterBlocks sum {sum(r['culled']['A'].get('rasterBlocks') or 0 for r in q)}, commands sum {sum(r['culled']['A'].get('commands') or 0 for r in q)}")
  area=[r['culled']['A']['blockArea']/(r['w']*r['h']) for r in q]; print(f"  block area share median {100*st.median(area):.1f}% mean {100*st.mean(area):.1f}%; units with block area>=90%: {sum(1 for a in area if a>=0.9)}")
print('\nTop 10 by lean saving (bytes):')
q=[r for r in R if r['culled']['A']['qualifies']]
for r in sorted(q,key=lambda r:r['finalRaster']-r['culled']['A']['lean'],reverse=True)[:10]:
  c=r['culled']['A'];print(f"| {r['file']} | {r['unit']} | {r['scale']} | {r['w']}x{r['h']} | {r['finalRaster']} | {c['lean']} | {r['finalRaster']-c['lean']} | {pct(c['lean'],r['finalRaster']):.1f}% | {c['total']} | {r['finalNoVisibleMark']} | {r['relaxed']['bytes']} |")
print('\nTop 10 by strict saving (bytes):')
for r in sorted(q,key=lambda r:r['finalRaster']-r['culled']['A']['total'],reverse=True)[:10]:
  c=r['culled']['A'];print(f"| {r['file']} | {r['unit']} | {r['scale']} | {r['finalRaster']} | {c['total']} | {r['finalRaster']-c['total']} | {pct(c['total'],r['finalRaster']):.1f}% |")
print('relaxed fallbacks', __import__('collections').Counter(r['relaxed'].get('fallback') for r in R))
print('final unmarked', __import__('collections').Counter(r.get('finalUnmarked') for r in R))
print('ms total',sum(r['ms'] for r in R))
