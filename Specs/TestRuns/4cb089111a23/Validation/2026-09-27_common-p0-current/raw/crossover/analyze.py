import hashlib
import json
import math
import statistics
from pathlib import Path

root = Path(__file__).parent
plan = json.loads((root/'plan.json').read_text())
invocations = json.loads((root/'invocations.json').read_text(encoding='utf-8-sig'))
assert len(invocations) == 10 and all(x['exit'] == 0 for x in invocations)
rows = []
identity = None
for entry in plan:
    d = Path(entry['directory'])
    files = json.loads((d/'common_smoke-runner-copy-hashes.json').read_text(encoding='utf-8-sig'))
    archive = d/'common_smoke-runner-artifacts'
    for item in files:
        assert hashlib.sha256((archive/item['path']).read_bytes()).hexdigest() == item['sha256']
    raw = archive/'artifacts/selftest/last_run/perf/perf_metrics.jsonl'
    values = [json.loads(s) for s in raw.read_text(encoding='utf-8-sig').splitlines()]
    values = [v for v in values if v['metric'].startswith('FileOps.CommonPresentation.')]
    assert len(values) == 97 and all(v['hr'] == 0 for v in values)
    ident = [{k:v[k] for k in ('metric','detail','durationUs','value0','value1')}
             for v in values if v['metric'].endswith(('ClientPixels','WorkloadFNV64'))]
    if identity is None: identity = ident
    assert identity == ident
    stages = {}
    timings = {}
    for value in values:
        metric = value['metric'].removeprefix('FileOps.CommonPresentation.')
        phase = value['detail']
        if metric in ('Memory','Cpu','Handles','GuiAndAttachments'):
            v = stages.setdefault(phase,{})
            if metric == 'Memory': v.update(private_bytes=value['value0'],working_set_bytes=value['value1'])
            if metric == 'Cpu': v.update(ui_cpu_us=value['value0'],process_cpu_us=value['value1'])
            if metric == 'Handles': v.update(handles=value['value0'],hosts=value['value1'])
            if metric == 'GuiAndAttachments': v.update(attachments=value['durationUs'],gdi=value['value0'],user=value['value1'])
        elif metric.endswith('Us'):
            timings.setdefault(metric,[]).append(value['durationUs'])
    assert len(stages) == 13
    cold = stages['cold']
    deltas = {phase:{key: value-cold[key] for key,value in data.items()} for phase,data in stages.items()}
    for phase in ('popup-destroyed','issues-destroyed','issues-cleaned'):
        assert deltas[phase]['hosts'] == 0 and deltas[phase]['attachments'] == 0
    rows.append(dict(ordinal=entry['ordinal'],policy=entry['policy'],verified_files=len(files),
                     run_id=json.loads((archive/'artifacts/selftest/last_run/run-all-tests-results.json').read_text(encoding='utf-8-sig'))['run_id'],
                     raw_metric_sha256=hashlib.sha256(raw.read_bytes()).hexdigest(),stages=stages,deltas_from_cold=deltas,timing_samples=timings))

def summary(values):
    return dict(samples=len(values),minimum=min(values),median=statistics.median(values),
                p95=sorted(values)[math.ceil(.95*len(values))-1],maximum=max(values))

comparison = {'identity':identity,'runs':rows,'timing':{},'stage_deltas':{}}
for metric in rows[0]['timing_samples']:
    groups={p:[v for row in rows if row['policy']==p for v in row['timing_samples'][metric]] for p in ('A','B')}
    comparison['timing'][metric] = {p:summary(v) for p,v in groups.items()}
    comparison['timing'][metric]['per_process_medians'] = [{p:[statistics.median(r['timing_samples'][metric]) for r in rows if r['policy']==p]} for p in ('A','B')]
for phase in rows[0]['stages']:
    comparison['stage_deltas'][phase]={key:{p:summary([r['deltas_from_cold'][phase][key] for r in rows if r['policy']==p]) for p in ('A','B')} for key in rows[0]['stages'][phase]}
(root/'analysis.json').write_bytes((json.dumps(comparison,indent=2)+'\n').encode('utf-8'))
print(json.dumps(dict(runs=len(rows),verified_files=sum(r['verified_files'] for r in rows),identity=identity,timing=comparison['timing']),indent=2))
for phase in ('popup-visible-before','popup-after-churn','popup-hidden-end','popup-destroyed','issues-selected-before','issues-hidden-end','issues-cleaned'):
    print(phase, {k:{p:v[p]['median'] for p in ('A','B')} for k,v in comparison['stage_deltas'][phase].items()})
