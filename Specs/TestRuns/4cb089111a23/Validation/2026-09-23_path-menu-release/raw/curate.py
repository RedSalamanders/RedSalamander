from pathlib import Path
import hashlib
import html
import shutil

ROOT = Path('Z:/src/RedSalamander')
EVIDENCE = Path('C:/RedSalamander.Perf/evidence/i26-ui/path-and-menu-20260923')
PACKET = ROOT / 'Specs/TestRuns/4cb089111a23/Validation/2026-09-23_path-menu-release'
GALLERY = ROOT / 'Specs/Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-path-menu-release'

def copy(src, dst):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dst)

def write_hashes(folder, files):
    rows = [f'{hashlib.sha256(p.read_bytes()).hexdigest()}  {p.relative_to(folder).as_posix()}\n' for p in sorted(files)]
    (folder / 'SHA256SUMS.txt').write_text(''.join(rows), encoding='utf-8')

for name in ('build-candidate.ps1', 'build-exit.txt', 'build-receipt.json', 'build.log',
             'candidate-inputs.json', 'capture-candidate.ps1', 'gallery-exit.txt', 'gallery-plan.json',
             'gallery.log', 'baseline-focused-metrics.jsonl', 'candidate-focused-metrics.jsonl',
             'compare-gallery.py', 'gallery-comparison.json', 'spec-inventory.log', 'curate.py'):
    copy(EVIDENCE / name, PACKET / 'raw' / name)
for name in ('results.json', 'run-all-tests-results.json', 'run-all-tests-dashboard.md',
             'run-all-tests-case-history.jsonl', 'trace.txt', 'fileops/results.json', 'fileops/trace.txt'):
    copy(EVIDENCE / 'runner-artifacts' / name, PACKET / 'raw/gallery-run' / name)
for label in ('fileops', 'pr'):
    archive = EVIDENCE / f'{label}-runner-artifacts/artifacts/selftest/last_run'
    if archive.is_dir():
        for name in ('results.json', 'run-all-tests-results.json', 'run-all-tests-results-v1.json',
                     'validation-plan.json', 'validation-plan-decision.json', 'entry-fingerprints.json', 'run-all-tests-dashboard.md',
                     'run-all-tests-case-history.jsonl', 'trace.txt', 'fileops/results.json', 'fileops/trace.txt'):
            if (archive / name).is_file():
                copy(archive / name, PACKET / f'raw/{label}-run' / name)
        for name in (f'qualify-{label}.ps1', f'{label}-exit.txt', f'{label}-plan.json', f'{label}.log'):
            if (EVIDENCE / name).is_file():
                copy(EVIDENCE / name, PACKET / 'raw' / name)
copy(EVIDENCE / 'spec-inventory-final.log', PACKET / 'raw/spec-inventory-final.log')
for src in (EVIDENCE / 'runner-artifacts/fileops/visual-gallery').iterdir():
    if src.is_file():
        copy(src, GALLERY / 'candidate' / src.name)
for name in ('candidate-inputs.json', 'gallery-comparison.json'):
    copy(EVIDENCE / name, GALLERY / name)
(PACKET / '.gitattributes').write_text('raw/** -text -whitespace\nSHA256SUMS.txt -text -whitespace\n', encoding='utf-8')
(GALLERY / '.gitattributes').write_text('candidate/** -text -whitespace\n*.json -text -whitespace\nSHA256SUMS.txt -text -whitespace\n', encoding='utf-8')
write_hashes(PACKET, (PACKET / 'raw').rglob('*.*'))
write_hashes(GALLERY, list((GALLERY / 'candidate').rglob('*.*')) + list(GALLERY.glob('*.json')))
cards = []
for path in sorted((GALLERY / 'candidate').glob('*.png')):
    name = html.escape(path.name)
    rejected = path.name.startswith('102-confirmation-')
    title = ('REJECTED: black confirmation capture — ' if rejected else '') + name
    cards.append(f'<article data-name="{name.lower()}"><h2>{title}</h2><a href="candidate/{name}"><img loading="lazy" src="candidate/{name}" alt="{title}"></a></article>')
page = '''<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>File Operations — path and destination-menu captures</title>
<style>:root{color-scheme:dark;font:16px/1.5 system-ui;background:#14171b;color:#eef1f4}body{max-width:1680px;margin:auto;padding:24px}a{color:#a6d5ff}h1{font-size:28px}h2{font-size:14px;overflow-wrap:anywhere}header{max-width:1050px}.filters{position:sticky;top:0;background:#14171bf5;padding:12px 0;z-index:1}input{font:inherit;padding:8px;width:min(90%,520px)}main{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,450px),1fr));gap:20px}article{background:#21262d;border:1px solid #535c66;border-radius:8px;padding:12px}article[hidden]{display:none}img{width:100%;height:420px;object-fit:contain;background:#080a0d}:focus-visible{outline:3px solid #8dd2ff;outline-offset:3px}</style>
<header><h1>Long filenames and destination menus</h1><p>Actual French application captures from the deterministic, nonactivating Release harness on physical 96/144-DPI monitors. All 30 targeted layout checks pass. <strong>The overall gallery fails at the black confirmation witness; later supplemental captures are missing.</strong> These are not mockups or final integrated acceptance.</p><p><a href="README.md">Review and limitations</a> · <a href="candidate/scenarios.tsv">Scenario manifest</a> · <a href="candidate/displays.tsv">Displays</a> · <a href="gallery-comparison.json">Paired comparison</a> · <a href="SHA256SUMS.txt">Original hashes</a></p></header>
<div class="filters"><label for="search">Filter by scenario, width or DPI</label><br><input id="search" type="search" placeholder="118-progress, 480dip, 96dpi…"><p id="count" aria-live="polite"></p></div><main>'''
page += '\n'.join(cards)
page += '''</main><script>const cards=[...document.querySelectorAll('article')],input=document.getElementById('search');function filter(){let count=0;for(const card of cards){card.hidden=!card.dataset.name.includes(input.value.toLowerCase());if(!card.hidden)count++;}document.getElementById('count').textContent=`${count} of ${cards.length} original captures`;}input.addEventListener('input',filter);filter();</script></html>'''
(GALLERY / 'gallery.html').write_text(page, encoding='utf-8')
print(f'Curated {len(list((GALLERY / "candidate").glob("*.png")))} original captures; raw packet {sum(p.stat().st_size for p in (PACKET / "raw").rglob("*.*"))} bytes.')
