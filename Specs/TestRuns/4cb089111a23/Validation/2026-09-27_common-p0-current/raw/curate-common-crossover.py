import hashlib
import json
import shutil
import subprocess
from pathlib import Path

root=Path(__file__).parent
runroot=root/'common-p0-current-crossover-20260927'
packet=Path('Z:/src/RedSalamander/Specs/TestRuns/4cb089111a23/Validation/2026-09-27_common-p0-current')
assert not (packet/'SHA256SUMS.txt').exists(), 'Do not overwrite completed evidence'
packet.mkdir(parents=True,exist_ok=True)
mapping=[]
def digest(data): return hashlib.sha256(data).hexdigest()
def copy(source,relative):
    data=source.read_bytes()
    target=packet/relative
    target.parent.mkdir(parents=True,exist_ok=True)
    target.write_bytes(data)
    mapping.append(dict(source=str(source),path=relative,sha256=digest(data),bytes=len(data)))
def selected_metrics(source,relative):
    data=source.read_bytes()
    selected=[]; numbers=[]
    for n,line in enumerate(data.splitlines(keepends=True),1):
        if json.loads(line).get('metric','').startswith('FileOps.CommonPresentation.'):
            selected.append(line);numbers.append(n)
    assert len(selected)==97
    target=packet/relative
    target.parent.mkdir(parents=True,exist_ok=True)
    output=b''.join(selected);target.write_bytes(output)
    mapping.append(dict(source=str(source),path=relative,sha256=digest(output),bytes=len(output),
                        selection='raw matching lines in original order',line_numbers=numbers,source_sha256=digest(data),source_bytes=len(data)))

analysis=json.loads((runroot/'analysis.json').read_text())
assert len(analysis['runs'])==10
for name in ('plan.json','invocations.json','analysis.json','machine.json','power-scheme.txt','run-crossover.ps1','analyze.py','README.md'):
    copy(runroot/name,'raw/crossover/'+name+('.receipt.txt' if name.endswith('.json') else ''))
copy(Path(__file__),'raw/curate-common-crossover.py')
copy(root/'prepare-common-crossover-20260927.py','raw/prepare-common-crossover.py')
for row in json.loads((runroot/'plan.json').read_text()):
    source=Path(row['directory']); prefix='raw/crossover/'+source.name
    for name in ('candidate-inputs.json','common_smoke-plan.json','common_smoke-runner-copy-hashes.json','invocations.json','processes-before.json','processes-after.json','common_smoke.log','driver.log','run.ps1'):
        if name.startswith('processes-') and not (source/name).exists():
            mapping.append(dict(source=str(source/name),omission='PowerShell emitted no file for the empty process pipeline; successful driver continued after its busy check'))
            continue
        copy(source/name,prefix+'/'+name+('.receipt.txt' if name.endswith('.json') else ''))
    last=source/'common_smoke-runner-artifacts/artifacts/selftest/last_run'
    for rel in ('results.json','run-all-tests-results.json','workspace-final.json','fileops/results.json','fileops/trace.txt'):
        copy(last/rel,prefix+'/'+rel+('.receipt.txt' if rel.endswith('.json') else ''))
    selected_metrics(last/'perf/perf_metrics.jsonl',prefix+'/perf_metrics.selected.jsonl.receipt.txt')
for version in ('common-p0-baseline-v1-20260927','common-p0-baseline-v2-20260927'):
    source=root/version
    for name in ('build.log','build-exit.txt','build-receipt.json','candidate-inputs.json','build-and-test.ps1','run-focused.ps1'):
        if (source/name).exists(): copy(source/name,'raw/'+version+'/'+name+('.receipt.txt' if name.endswith('.json') else ''))
source=root/'exact-bound-results-candidate-v5-20260927'
copy(source/'build-receipt.json','raw/current-v5-build-receipt.json.receipt.txt')
for repo,relative in [('C:/RsI26P0','p0-test-overlays.patch'),('C:/RsI26Final','current-overlays.patch')]:
    data=subprocess.check_output(['git','diff','--binary'],cwd=repo)
    (packet/'raw'/relative).write_bytes(data)
fixture=Path('C:/RsI26P0/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.CommonPresentationPerf.cpp')
assert digest(fixture.read_bytes())=='9c817b6cfe7570a3fea9c2d0b5a95da9379a5368df7612debd80628f095224fd'
copy(fixture,'raw/CommonPresentationPerf.cpp.source.txt')
(packet/'raw/file-map.receipt.txt').write_bytes((json.dumps(mapping,indent=2)+'\n').encode('utf-8'))
(packet/'.gitattributes').write_bytes(b'raw/** -text -whitespace\nSHA256SUMS.txt -text -whitespace\n')
(packet/'README.md').write_bytes('''# Original P0 versus current File Operations presentation

Ten fresh x64 Release processes pass (five per revision), in the predeclared ABBAABBAAB order, with zero disk-audit findings. All 270 copied runner files were independently rehashed. The test fixture is byte-identical, SHA256 `9c817b6cfe7570a3fea9c2d0b5a95da9379a5368df7612debd80628f095224fd`. Every run uses French/dark/reduced-motion/default-density, eight running Copy/Move cards and 128 selected Issues rows, equal workload digest `14201940575619722830`, and 1140 x 720 client pixels at 144 DPI for both windows. Power, hardware, font, process, executable and source identities are retained.

A is original P0 `e61e349c3e8d133e9d52a4c63ba4c49266bc4db5` with DxUi `f5c7fc7403e352425b8e82b7a4da1c042d942ede`. Only test registration, its private helper include, existing-render-count accessor and common fixture are overlaid. Its first build failed on the missing test-only include; that failed log is retained. Corrected full-solution build receipt is `f510718fd9d42c265bcc6317b12012088f266f7d4776a0e2930b4827bdeca7ed`.

B is the frozen V5 consumer on isolated base `ccebd5af68eb45f2e9deed64a7063780e71eb1ce`, 36 overlays matching primary `a86dd111`, pin `78b3de389a189c7f86f611787e0489fb6d474218`. Full-solution receipt is `e231b251d2a6a43e2e209b45aae9b2959b8b969727caeda6e693040f4d34f2da`. The current grid/menu branches are not adopted. This comparison includes the whole intervening consumer/library redesign; it cannot attribute a difference to one control or optimization.

| Synchronous work (eight redraws/updates per sample) | P0 median / p95 (ms) | Current median / p95 (ms) | Samples per revision |
| --- | ---: | ---: | ---: |
| Running popup redraw | 66.711 / 85.529 | 66.652 / 66.844 | 25 |
| Eight-task progress churn | 66.556 / 68.502 | 66.661 / 66.820 | 160 |
| Selected Issues redraw | 66.649 / 150.113 | 66.684 / 150.062 | 25 |

P95 uses pooled nearest-rank samples; the raw report also retains each process median. These are synchronous wall times, not presented FPS or frame intervals. The median differences are small; this fixture does not show a material redraw slowdown. Sparse tail samples and display synchronization limit stronger timing claims. No threshold was invented or relaxed.

Median process-private growth from each process's pre-popup sample is 44.445 MiB for P0 versus 8.227 MiB current when the popup first settles. After the full popup/Issues lifetime, it is 29.945 MiB versus -1.660 MiB. Negative process growth means other initialization/allocator memory was released during the scenario; it is not negative UI allocation. Actual current private memory is lower in every compared process range: first-open 263.66–269.37 MiB versus P0 299.53–304.04 MiB. Full per-stage absolute and relative values are retained. Process CPU through teardown has median 1,421,875 us P0 versus 1,046,875 us current; the Windows process/thread accounting resolution is coarse.

All ten runs verify zero hosted rendering while hidden, restored selected row identity, removed diagnostic history and return of host/attachment counts after each destruction. Remaining process handles/GDI/USER counts are recorded rather than assumed to be zero. This is bounded lifetime evidence, not a long-duration leak proof.

The original-P0 common-workload comparison is complete. It does not qualify surface bytes, final combined pin, newer selected-detail/graph-disclosure parity, native monitor movement, keyboard/UIA/AT, or the complete P5 resource matrix. Equal-feature comparisons and those required gates remain open. The accepted grid memory tradeoff is not applied to this consumer measurement.

Full runner trees and complete metric streams remain in `C:/RedSalamander.Perf/evidence/i26-ui/common-p0-current-crossover-20260927`. This packet retains all outcomes, source overlays, build receipts and exact raw common-metric lines with source hashes/line numbers. Empty process snapshots produced no file through PowerShell's empty pipeline; the file map explicitly records those omissions. Before/after snapshots cannot exclude activity occurring only between them. SHA256SUMS binds every payload. Baseline smoke receipt and original failed setup stay in their separate external directories; the ten-run comparison starts only after both fixture smokes pass.
'''.encode('utf-8'))
files=sorted(p for p in packet.rglob('*') if p.is_file())
assert max(p.stat().st_size for p in files)<2*1024*1024
assert sum(p.stat().st_size for p in files)<5*1024*1024
(packet/'SHA256SUMS.txt').write_bytes(''.join(digest(p.read_bytes())+'  '+p.relative_to(packet).as_posix()+'\n' for p in files).encode('ascii'))
print(f'{len(files)} payloads; {sum(p.stat().st_size for p in files)} bytes')
