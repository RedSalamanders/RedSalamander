from pathlib import Path
import hashlib,json,re,statistics,collections
from PIL import Image,ImageChops,ImageDraw
root=Path('C:/RedSalamander.Perf/evidence/i26-ui')
base=root/'speed-selector-baseline-20260923'; cand=root/'speed-selector-qualified-20260923'
results={}; comparisons=[]
for label,e in [('baseline',base),('candidate',cand)]:
 run=e/'runner-artifacts/artifacts/selftest/last_run'; data=(run/'fileops/trace.txt').read_bytes();text=data.decode('utf-16' if data[:2] in (b'\xff\xfe',b'\xfe\xff') else 'utf-8-sig')
 labels=[l for l in text.splitlines() if 'SpeedLimitLabelFit:' in l];changes=[l for l in text.splitlines() if 'SpeedLimitValueChange:' in l]
 samples=collections.defaultdict(list)
 for dpi,w,r,n,t in re.findall(r'SpeedLimitRedraw: dpi=(\d+) width=(\d+) round=(\d+) frames=(\d+) elapsedUs=(\d+)',text):samples[f'{dpi}dpi-{w}dip'].append(int(t)/int(n))
 results[label]={'checks':len(labels),'failures':sum('fit=false' in l for l in labels),'sameWindowChanges':len(changes),'sameWindowFailures':sum('fit=false' in l for l in changes),'samplesUsPerSynchronousRedraw':samples,'medianUsPerSynchronousRedraw':{k:statistics.median(v) for k,v in samples.items()}}
 if label=='candidate':assert results[label]['failures']==0 and results[label]['sameWindowFailures']==0
 gallery=run/'fileops/visual-gallery'; results[label]['captures']=len(list(gallery.glob('*.png')))
bg=base/'runner-artifacts/artifacts/selftest/last_run/fileops/visual-gallery';cg=cand/'runner-artifacts/artifacts/selftest/last_run/fileops/visual-gallery'
assert {p.name for p in bg.glob('*.png')}=={p.name for p in cg.glob('*.png')}
for p in sorted(bg.glob('*.png')):
 q=cg/p.name;a=Image.open(p).convert('RGBA');b=Image.open(q).convert('RGBA')
 bounds=ImageChops.difference(a,b).getbbox(alpha_only=False) if a.size==b.size else (0,0,max(a.width,b.width),max(a.height,b.height))
 comparisons.append({'file':p.name,'baselineSha256':hashlib.sha256(p.read_bytes()).hexdigest(),'candidateSha256':hashlib.sha256(q.read_bytes()).hexdigest(),'baselineDimensions':list(a.size),'candidateDimensions':list(b.size),'changedBounds':list(bounds) if bounds else None})
results['sameFixture']=json.loads((base/'candidate-inputs.json').read_text(encoding='utf-8-sig'))[-1:] # full binding checked separately below
bf={x['path']:x['sha256'] for x in json.loads((base/'candidate-inputs.json').read_text(encoding='utf-8-sig'))};cf={x['path']:x['sha256'] for x in json.loads((cand/'candidate-inputs.json').read_text(encoding='utf-8-sig'))}
changedSources=[p for p in bf if bf[p]!=cf[p]];assert set(changedSources)=={'RedSalamander/FolderWindow.FileOperations.Popup.cpp','RedSalamander/FolderWindow.FileOperations.Popup.h'}
results['changedSourceFiles']=changedSources;results['unchangedSourceHashes']={p:sha for p,sha in bf.items() if p not in changedSources};results.pop('sameFixture')
results['redrawMedianChangePercent']={k:100*(v/results['baseline']['medianUsPerSynchronousRedraw'][k]-1) for k,v in results['candidate']['medianUsPerSynchronousRedraw'].items()}
results['unchangedImages']=sum(x['changedBounds'] is None for x in comparisons);results['changedImages']=len(comparisons)-results['unchangedImages'];results['scope']='Synchronous redraw samples only; not presented FPS, CPU/idle, allocation, or resource acceptance. All raw samples retained.'
(cand/'speed-comparison.json').write_text(json.dumps(results,indent=2)+'\n',encoding='utf-8');(cand/'capture-comparison.json').write_text(json.dumps(comparisons,indent=2)+'\n',encoding='utf-8')
selected=[x for x in comparisons if x['changedBounds'] is not None];sheet=Image.new('RGB',(4*440,((len(selected)+3)//4)*400),'#333');draw=ImageDraw.Draw(sheet)
for i,x in enumerate(selected):
 im=Image.open(cg/x['file']).convert('RGB');im.thumbnail((432,372));left=i%4*440;top=i//4*400;sheet.paste(im,(left+(440-im.width)//2,top+22));draw.text((left+4,top+3),x['file'][:53],fill='white')
if selected:sheet.save(cand/'changed-captures-contact.jpg')
print(json.dumps({k:v for k,v in results.items() if k!='unchangedSourceHashes'},indent=2));print('Changed:', '\n'.join(x['file'] for x in selected))
