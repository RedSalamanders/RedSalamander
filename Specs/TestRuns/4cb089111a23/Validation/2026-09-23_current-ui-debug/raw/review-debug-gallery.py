from pathlib import Path
import hashlib,json,re
from PIL import Image,ImageOps,ImageDraw,ImageChops
e=Path('C:/RedSalamander.Perf/evidence/i26-ui/waiting-issues-debug-20260923')
run=e/'runner-artifacts/artifacts/selftest/last_run'
gallery=run/'fileops/visual-gallery'
release=e.parent/'waiting-issues-isolated-20260923/runner-artifacts/artifacts/selftest/last_run/fileops/visual-gallery'
result=json.loads((run/'run-all-tests-results.json').read_text(encoding='utf-8-sig'))
assert result['exit_code']==0 and result['test_sandbox_audit']['issue_count']==0
trace_bytes=(run/'fileops/trace.txt').read_bytes()
trace=trace_bytes.decode('utf-16' if trace_bytes[:2] in (b'\xff\xfe',b'\xfe\xff') else 'utf-8-sig')
transitions=re.findall(r'WaitingCardTransition: dpi=(\d+) width=(\d+) phase=(\d+) frames=(\d+) fit=(\w+) settled=(\w+)',trace)
assert len(transitions)==48 and all(t[4:]==('true','true') and int(t[3])>=2 for t in transitions)
records=[]
for p in sorted(gallery.glob('*.png')):
 old=release/p.name
 digest=hashlib.sha256(p.read_bytes()).hexdigest()
 previous=hashlib.sha256(old.read_bytes()).hexdigest() if old.exists() else None
 pixel_bounds=None
 if digest!=previous:
  with Image.open(old) as a,Image.open(p) as b:
   assert a.size==b.size
   pixel_bounds=ImageChops.difference(a.convert('RGBA'),b.convert('RGBA')).getbbox(alpha_only=False)
 records.append({'file':p.name,'sha256':digest,'release_sha256':previous,'identical':digest==previous,'pixelDifferenceBounds':pixel_bounds,'onlyTitleIconDiffers':pixel_bounds in ((13,10,37,34),(9,7,25,23))})
assert len(records)==424
(e/'capture-comparison.json').write_text(json.dumps(records,indent=2)+'\n',encoding='utf-8')
changed=[r for r in records if not r['identical']]
summary={'total_captures':len(records),'byte_identical_to_release':len(records)-len(changed),'changed':len(changed),'transition_observations':len(transitions),'minimum_transition_frames':min(int(t[3]) for t in transitions),'native_dpi':sorted({int(t[0]) for t in transitions}),'widths':sorted({int(t[1]) for t in transitions}),'result':result}
summary['only_title_icon_changes']=sum(r['onlyTitleIconDiffers'] for r in records)
summary['all_body_pixels_equal']=all(r['identical'] or r['onlyTitleIconDiffers'] for r in records)
(e/'gallery-review-summary.json').write_text(json.dumps(summary,indent=2)+'\n',encoding='utf-8')
review=e/'review-sheets';review.mkdir(exist_ok=True)
for offset in range(0,len(changed),4):
 sheet=Image.new('RGB',(1600,1800),'#202020');draw=ImageDraw.Draw(sheet)
 for i,record in enumerate(changed[offset:offset+4]):
  with Image.open(gallery/record['file']) as source:
   tile=ImageOps.contain(source.convert('RGB'),(780,820))
  x=(i%2)*800;y=(i//2)*900
  sheet.paste(tile,(x+(800-tile.width)//2,y+60))
  draw.text((x+8,y+8),record['file'][:85],fill='white')
  draw.text((x+8,y+28),record['file'][85:],fill='white')
 sheet.save(review/f'{offset//4+1:02}.jpg',quality=92)
print(json.dumps({k:v for k,v in summary.items() if k!='result'},indent=2))
