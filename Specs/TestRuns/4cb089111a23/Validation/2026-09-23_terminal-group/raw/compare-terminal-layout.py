from pathlib import Path
import json,re,statistics,collections,hashlib
from PIL import Image,ImageChops
e=Path('C:/RedSalamander.Perf/evidence/i26-ui')
dirs={'baseline':e/'terminal-group-baseline-v2-20260923','candidate':e/'terminal-group-qualified-20260923'}
summary={}; galleries={}
for label,folder in dirs.items():
    run=folder/'runner-artifacts/artifacts/selftest/last_run'
    raw=(run/'fileops/trace.txt').read_bytes();text=raw.decode('utf-16' if raw[:2] in (b'\xff\xfe',b'\xfe\xff') else 'utf-8-sig')
    records={name:[line for line in text.splitlines() if name+':' in line] for name in
             ('CompletedGroupLabelFit','TerminalSummaryFit','TerminalSummaryChange','SpeedLimitLabelFit')}
    samples=collections.defaultdict(list)
    pattern=r'TerminalLayoutRedraw: (\S+) dpi=(\d+) width=(\d+) round=(\d+) frames=(\d+) elapsedUs=(\d+) privateBytes=(\d+) workingSetBytes=(\d+) gdi=(\d+) user=(\d+)'
    pending=[]
    for line in text.splitlines():
        match=re.search(pattern,line)
        if match:
            scene,dpi,width,round_,frames,us,private,ws,gdi,user=match.groups()
            pending.append(dict(scene=scene,dpi=int(dpi),width=int(width),round=int(round_),usPerRedraw=int(us)/int(frames),privateBytes=int(private),workingSetBytes=int(ws),gdi=int(gdi),user=int(user)))
        capture=re.search(r'VisualGallery: (\S+\.png)',line)
        if capture and pending:
            filename=capture.group(1)
            assert len(pending)==5 and [v['round'] for v in pending]==list(range(5)), filename
            for value in pending:
                assert filename.startswith(value['scene']+'__') and f"__{value['width']}dip__" in filename and filename.endswith(f"__{value['dpi']}dpi.png"), filename
            assert filename not in samples, filename
            samples[filename]=pending;pending=[]
    assert not pending
    summary[label]={'checks':{name:dict(count=len(lines),failures=sum('fit=false' in line for line in lines)) for name,lines in records.items()},'samples':samples}
    if label=='candidate':assert all(v['count']>0 and v['failures']==0 for v in summary[label]['checks'].values())
    galleries[label]=run/'fileops/visual-gallery'
assert set(summary['baseline']['samples'])==set(summary['candidate']['samples'])
paired={}
for name,a in summary['baseline']['samples'].items():
    b=summary['candidate']['samples'][name];assert len(a)==len(b)==5
    paired[name]={}
    for metric in ('usPerRedraw','privateBytes','workingSetBytes','gdi','user'):
        av=statistics.median(v[metric] for v in a);bv=statistics.median(v[metric] for v in b)
        paired[name][metric]=dict(baseline=av,candidate=bv,delta=bv-av,changePercent=100*(bv/av-1) if av else None)
summary['pairedMedians']=paired
summary['scope']='Same-machine in-process synchronous redraw and process observations; not presented FPS, live heap attribution, GPU surface bytes or final P5 qualification.'
assert {p.name for p in galleries['baseline'].glob('*.png')}=={p.name for p in galleries['candidate'].glob('*.png')}
images=[]
for p in sorted(galleries['baseline'].glob('*.png')):
    q=galleries['candidate']/p.name;a=Image.open(p).convert('RGBA');b=Image.open(q).convert('RGBA')
    box=ImageChops.difference(a,b).getbbox(alpha_only=False) if a.size==b.size else (0,0,max(a.width,b.width),max(a.height,b.height))
    images.append(dict(file=p.name,baselineSha256=hashlib.sha256(p.read_bytes()).hexdigest(),candidateSha256=hashlib.sha256(q.read_bytes()).hexdigest(),baselineDimensions=list(a.size),candidateDimensions=list(b.size),changedBounds=list(box) if box else None))
summary['captures']=len(images);summary['changedImages']=sum(v['changedBounds'] is not None for v in images)
(dirs['candidate']/'layout-comparison.json').write_text(json.dumps(summary,indent=2)+'\n',encoding='utf-8')
(dirs['candidate']/'capture-comparison.json').write_text(json.dumps(images,indent=2)+'\n',encoding='utf-8')
print(json.dumps({label:summary[label]['checks'] for label in dirs},indent=2));print('Captures:',len(images),'changed:',summary['changedImages'])
