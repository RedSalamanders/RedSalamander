import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import {
  Folder20Regular, Document20Regular, ArrowLeft20Regular, ArrowRight20Regular,
  ArrowUp20Regular, ChevronRight16Regular, ChevronDown16Regular, PanelRight20Regular,
  PanelBottom20Regular, FullScreenMaximize20Regular, Search20Regular, Dismiss20Regular,
  MoreHorizontal20Regular, Copy20Regular, EyeOff20Regular, Eye20Regular,
  ArrowDownload20Regular, Delete20Regular, Checkmark20Regular, Warning20Regular,
  Info20Regular, ArrowClockwise20Regular, Settings20Regular, WindowNew20Regular,
  DataTreemap24Regular, Color20Regular, CheckmarkCircle20Regular, Filter20Regular,
  WeatherMoon20Regular, WeatherSunny20Regular,
} from '@fluentui/react-icons';
import { fixture, nodes, categories, groupColors, isInside, isExcluded, amount,
  unknownDisk, formatted, metricText, breakdown, hasType, treemap } from './model.js';

const empty = new Set();
const metricNames = { size: 'Size', count: 'File count', disk: 'Size on disk' };
const colorNames = { folder: 'Folder groups', type: 'File type', age: 'Age · modified', newest: 'Age · newest inside' };
const ageColors = ['#89b8ac', '#a9bf88', '#dec080', '#d19e8e'];
const folderCount = [...nodes.values()].filter(n => n.kind === 'folder').length - 1;
const ageBand = date => !date ? -1 : date >= '2026-10-15' ? 0 : date >= '2026-09-22' ? 1 : date >= '2025-10-22' ? 2 : 3;
const parentName = node => nodes.get(node.parent)?.path || 'D:\\Workspaces';
const pct = (value, total) => total ? `${(100 * value / total).toLocaleString('en-US', { maximumFractionDigits: 1 })}%` : '—';
function IconButton({ icon: Icon, title, children, ...props }) {
  return <button className="icon-button" title={title} aria-label={title} {...props}><Icon aria-hidden="true" />{children}</button>;
}
function ItemIcon({ node, excluded }) {
  const Icon = node.kind === 'folder' ? Folder20Regular : node.kind === 'aggregate' ? MoreHorizontal20Regular : Document20Regular;
  return <span className={`item-icon ${node.kind}`} role="img" aria-label={node.kind === 'aggregate' ? 'Summarized entries' : node.kind === 'folder' ? 'Folder' : 'File'}><Icon aria-hidden="true" />{excluded && <EyeOff20Regular className="exclude-overlay" aria-hidden="true" />}</span>;
}
function Modal({ title, children, onClose, className = '' }) {
  const ref = useRef(null), closeRef = useRef(onClose); closeRef.current = onClose;
  useEffect(() => {
    const previous = document.activeElement;
    ref.current?.querySelector('button, input, select, textarea')?.focus();
    const key = event => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); closeRef.current(); }
      if (event.key === 'Tab') {
        const all = [...ref.current.querySelectorAll('button:not(:disabled),input:not(:disabled),select,textarea,[tabindex="0"]')];
        if (event.shiftKey && document.activeElement === all[0]) { event.preventDefault(); all.at(-1)?.focus(); }
        else if (!event.shiftKey && document.activeElement === all.at(-1)) { event.preventDefault(); all[0]?.focus(); }
      }
    };
    document.addEventListener('keydown', key, true);
    return () => { document.removeEventListener('keydown', key, true); previous?.focus(); };
  }, []);
  return <div className="modal-scrim" onMouseDown={event => event.target === event.currentTarget && onClose()}>
    <section className={`dialog ${className}`} role="dialog" aria-modal="true" aria-label={title} ref={ref}>
      <header><h2>{title}</h2><IconButton icon={Dismiss20Regular} title="Close dialog" onClick={onClose} /></header>{children}
    </section>
  </div>;
}

function MapView({ current, metric, colorMode, exclusions, removed, selected, marks, matches, onSelect, onEnter, onContext }) {
  const ref = useRef(null), [size, setSize] = useState({ w: 800, h: 600 });
  useLayoutEffect(() => {
    const observer = new ResizeObserver(([entry]) => setSize({ w: entry.contentRect.width, h: entry.contentRect.height }));
    observer.observe(ref.current); return () => observer.disconnect();
  }, []);
  const color = (node, group) => {
    if (colorMode === 'type') return categories[node.type]?.color || '#cad4cf';
    if (colorMode === 'age' || colorMode === 'newest') {
      const band = ageBand(colorMode === 'newest' && node.kind === 'folder' ? node.newest : node.modified);
      return band < 0 ? '#bdc5c9' : ageColors[band];
    }
    return node.kind === 'aggregate' ? '#c3ccc8' : groupColors[group % groupColors.length];
  };
  const render = (children, w, h, depth = 0, group = 0, branchMatch = true) => treemap(children, w, h, n => amount(n, metric, exclusions, removed)).map(rect => {
    const n = rect.node, g = depth ? group : children.indexOf(n), branchVisible = depth ? branchMatch : matches(n), nested = n.children?.length && rect.w > 115 && rect.h > 100 && depth < 2;
    const exact = metric !== 'disk' || !unknownDisk(n, exclusions, removed);
    const value = `${exact ? '' : '≥ '}${metricText(amount(n, metric, exclusions, removed), metric)}`;
    return <div key={n.id} className={`map-tile ${n.kind} ${nested ? 'nested' : ''} ${selected === n.id ? 'selected' : ''} ${!branchVisible && !depth ? 'dimmed' : ''}`}
      style={{ left: rect.x, top: rect.y, width: rect.w, height: rect.h, '--tile-color': color(n, g) }}
      onClick={event => { event.stopPropagation(); onSelect(n.id); }} onContextMenu={event => { event.preventDefault(); event.stopPropagation(); onContext(n, event.clientX, event.clientY); }}
      title={`${n.path}\n${value}${n.kind === 'aggregate' ? '\nSummarized detail; individual names are not retained.' : ''}`}>
      <button className="tile-hit" aria-label={`${n.kind === 'folder' ? 'Explore' : 'Select'} ${n.name}`} onClick={event => { event.stopPropagation(); onSelect(n.id); if (n.kind === 'folder') onEnter(n.id); }}>
        {rect.w > 48 && rect.h > 23 && <span className="tile-name"><ItemIcon node={n} /><span>{n.name}</span>{marks.has(n.id) && <Checkmark20Regular />}</span>}
        {!nested && rect.w > 72 && rect.h > 48 && <span className="tile-size">{value}</span>}
        {nested && rect.w > 170 && <span className="tile-heading-size">{value}</span>}
      </button>
      {nested && <div className="tile-children" style={{ top: 27, left: 4 }}>{render(n.children, rect.w - 8, rect.h - 31, depth + 1, g, branchVisible)}</div>}
    </div>;
  });
  const visibleAmount = amount(current, metric, exclusions, removed);
  return <div className="treemap" ref={ref} aria-label={`${metricNames[metric]} treemap for ${current.name}`}>
    {visibleAmount ? render(current.children || [current], size.w, size.h) : <div className="map-empty"><EyeOff20Regular /><strong>No measured items to display</strong><span>Restore excluded items or choose another size measure.</span></div>}
  </div>;
}

export function App() {
  const [history, setHistory] = useState([fixture.id]), [historyIndex, setHistoryIndex] = useState(0);
  const [selectedId, setSelectedId] = useState(fixture.children[0].id);
  const [exclusions, setExclusions] = useState(new Set()), [marks, setMarks] = useState(new Set());
  const [removed, setRemoved] = useState(new Set());
  const [metric, setMetric] = useState('size'), [colorMode, setColorMode] = useState('folder');
  const [layout, setLayout] = useState('right'), [narrow, setNarrow] = useState(window.innerWidth < 900);
  const [search, setSearch] = useState(''), [typeFilter, setTypeFilter] = useState(''), [ageFilter, setAgeFilter] = useState('');
  const [dateBasis, setDateBasis] = useState('modified'), [showTypes, setShowTypes] = useState(false), [details, setDetails] = useState(false);
  const [menu, setMenu] = useState(null), [modal, setModal] = useState(null), [context, setContext] = useState(null);
  const [scenario, setScenario] = useState('ready'), [dark, setDark] = useState(false), [toast, setToast] = useState('');
  const [closed, setClosed] = useState(false), [exportType, setExportType] = useState('json'), [deleteMode, setDeleteMode] = useState('recycle');
  const [sortBy, setSortBy] = useState('size'), [advanced, setAdvanced] = useState(false);
  const filterRef = useRef(null);
  useLayoutEffect(() => {
    if (!menu) return;
    const popup = document.querySelector('.menu-anchor .menu');
    if (!popup) return;
    const rect = popup.parentElement.getBoundingClientRect(), height = popup.offsetHeight;
    popup.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - popup.offsetWidth - 8))}px`;
    popup.style.right = 'auto'; popup.style.bottom = 'auto';
    popup.style.top = `${Math.max(8, rect.bottom + height + 8 > window.innerHeight ? rect.top - height - 4 : rect.bottom + 4)}px`;
  }, [menu, narrow, layout]);
  useEffect(() => {
    if ((!menu && !context) || menu === 'filters') return;
    const previous = document.activeElement, popup = document.querySelector(context ? '.context-menu' : '.menu-anchor .menu');
    const buttons = [...(popup?.querySelectorAll('button:not(:disabled)') || [])]; buttons[0]?.focus();
    const arrow = event => {
      if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault(); const index = buttons.indexOf(document.activeElement);
      buttons[event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length]?.focus();
    };
    document.addEventListener('keydown', arrow);
    return () => { document.removeEventListener('keydown', arrow); if (previous?.isConnected) previous.focus(); };
  }, [menu, context]);
  const current = nodes.get(history[historyIndex]), selected = nodes.get(selectedId);
  const scanning = scenario === 'scanning' || scenario === 'network';
  const effectiveLayout = layout === 'right' && narrow ? 'bottom' : layout;
  const visible = scenario === 'empty' ? 0 : amount(current, metric, exclusions, removed), discovered = scenario === 'empty' ? 0 : amount(fixture, 'size', empty, removed);
  const shownBytes = scenario === 'empty' ? 0 : amount(current, 'size', exclusions, removed), originalBytes = scenario === 'empty' ? 0 : amount(current, 'size', empty, removed);
  const activeFilters = !!(search || typeFilter || ageFilter);
  const dateFor = node => dateBasis === 'newest' && node.kind === 'folder' ? node.newest : node.modified;
  const matches = node => (!search || node.name.toLocaleLowerCase().includes(search.toLocaleLowerCase())) && hasType(node, typeFilter) &&
    (!ageFilter || (dateFor(node) && dateFor(node) < (ageFilter === 'year' ? '2025-10-22' : '2026-09-22')));
  const children = scenario === 'empty' ? [] : (current.children || []).filter(n => !removed.has(n.id));
  const rows = children.filter(matches).sort((a, b) => sortBy === 'name' ? a.name.localeCompare(b.name) : amount(b, metric, empty, removed) - amount(a, metric, empty, removed));
  const types = scenario === 'empty' ? [] : breakdown(current, exclusions, removed);
  const markedNodes = [...marks].map(id => nodes.get(id)).filter(n => n && !removed.has(n.id));
  const markedBytes = markedNodes.reduce((sum, n) => sum + amount(n, 'size', empty, removed), 0);
  const problemNode = fixture.children.find(n => n.issue);
  const clearFilters = () => { setSearch(''); setTypeFilter(''); setAgeFilter(''); };
  const announce = text => setToast(text);
  const navigate = id => {
    if (!id || !nodes.get(id)?.children) return;
    setHistory(prev => [...prev.slice(0, historyIndex + 1), id]); setHistoryIndex(historyIndex + 1);
    setSelectedId(null); clearFilters(); setContext(null); setMenu(null);
  };
  const goHistory = direction => { const next = historyIndex + direction; if (next >= 0 && next < history.length) { setHistoryIndex(next); clearFilters(); setSelectedId(null); } };
  const toggleExclusion = node => {
    if (node.kind === 'aggregate') return;
    const inherited = [...exclusions].find(id => id !== node.id && isInside(node.id, id));
    if (inherited) { announce(`Excluded by ${nodes.get(inherited).name}. Restore that folder or use Restore all.`); setContext(null); setMenu(null); return; }
    setExclusions(prev => { const next = new Set(prev); if (next.has(node.id)) next.delete(node.id); else { [...next].forEach(id => { if (isInside(id, node.id)) next.delete(id); }); next.add(node.id); } return next; });
    setMarks(prev => new Set([...prev].filter(id => !isInside(id, node.id) && !isInside(node.id, id)))); setContext(null); setMenu(null);
  };
  const toggleMark = node => {
    if (node.kind === 'aggregate' || isExcluded(node, exclusions) || node.incomplete) return;
    setMarks(prev => { const next = new Set(prev); if (next.has(node.id)) next.delete(node.id); else {
      if ([...next].some(id => isInside(node.id, id))) return next;
      [...next].forEach(id => { if (isInside(id, node.id)) next.delete(id); }); next.add(node.id);
    } return next; });
  };
  const markMatching = () => {
    const targets = rows.filter(n => n.kind !== 'aggregate' && !n.incomplete && !isExcluded(n, exclusions));
    setMarks(prev => { const next = new Set(prev); targets.forEach(n => { if (![...next].some(id => isInside(n.id, id))) {
      [...next].forEach(id => { if (isInside(id, n.id)) next.delete(id); }); next.add(n.id);
    } }); return next; });
    setMenu(null); announce(`${targets.length} matching items added to review. Nothing has been deleted.`);
  };
  const copyPath = async (node = selected || current) => {
    if (node.kind === 'aggregate') return;
    setMenu(null); setContext(null);
    try { await navigator.clipboard.writeText(node.path); announce('Path copied'); }
    catch { setModal({ kind: 'copy', node }); }
  };
  const contextFor = (node, x, y) => { setSelectedId(node.id); setMenu(null); setContext({ node, x: Math.min(x, window.innerWidth - 254), y: Math.min(y, window.innerHeight - 274) }); };
  const reset = () => {
    document.querySelector('.item-list')?.scrollTo(0, 0);
    setHistory([fixture.id]); setHistoryIndex(0); setSelectedId(fixture.children[0].id); setExclusions(new Set()); setMarks(new Set()); setRemoved(new Set());
    setMetric('size'); setColorMode('folder'); setLayout('right'); clearFilters(); setMenu(null); setModal(null); setContext(null); setScenario('ready'); setClosed(false); setShowTypes(false); setDetails(false);
  };
  useEffect(() => { if (!toast) return; const timeout = setTimeout(() => setToast(''), 4200); return () => clearTimeout(timeout); }, [toast]);
  useEffect(() => { document.querySelector('.item-list')?.scrollTo(0, 0); }, [historyIndex, search, typeFilter, ageFilter, sortBy]);
  useEffect(() => { const resize = () => setNarrow(window.innerWidth < 900); window.addEventListener('resize', resize); return () => window.removeEventListener('resize', resize); }, []);
  useEffect(() => {
    const key = event => {
      if (modal || event.defaultPrevented) return;
      const editing = /INPUT|TEXTAREA|SELECT/.test(event.target.tagName);
      if (event.key === 'Escape') {
        if (menu || context) { setMenu(null); setContext(null); }
        else if (activeFilters) clearFilters();
        else if (scanning) { setScenario('partial'); announce('Scan cancelled. Discovered results remain available.'); }
        else setClosed(true);
        event.preventDefault(); return;
      }
      if (editing) return;
      if (event.altKey && event.key === 'Insert') { event.preventDefault(); copyPath(); }
      if (event.ctrlKey && event.key === '1') { event.preventDefault(); setLayout('right'); }
      if (event.ctrlKey && event.key === '2') { event.preventDefault(); setLayout('bottom'); }
      if (event.key === 'Backspace') { event.preventDefault(); if (current.parent) navigate(current.parent); }
      if (event.altKey && event.key === 'ArrowLeft') { event.preventDefault(); goHistory(-1); }
      if (event.altKey && event.key === 'ArrowRight') { event.preventDefault(); goHistory(1); }
      if (event.ctrlKey && event.key.toLowerCase() === 'f') { event.preventDefault(); filterRef.current?.focus(); }
    };
    window.addEventListener('keydown', key); return () => window.removeEventListener('keydown', key);
  });
  const closeModal = () => setModal(null);
  const ancestors = []; let ancestor = current; while (ancestor) { ancestors.unshift(ancestor); ancestor = nodes.get(ancestor.parent); }
  const selectionExcluded = selected && isExcluded(selected, exclusions);
  const selectionValue = selected ? amount(selected, metric, selectionExcluded ? empty : exclusions, removed) : 0;
  const selectionUnknown = selected && metric === 'disk' && unknownDisk(selected, exclusions, removed);
  const menuAction = action => { setMenu(null); action(); };
  const choice = (label, value, actual, action, Icon) => <button role="menuitemradio" aria-checked={value === actual} onClick={() => menuAction(() => action(value))}>
    {Icon && <Icon />}<span>{label}</span>{value === actual && <Checkmark20Regular className="menu-check" />}</button>;
  const itemActions = node => <>
    {node.kind === 'folder' && <button onClick={() => navigate(node.id)}><Folder20Regular /><span>Explore folder</span><kbd>Enter</kbd></button>}
    <button disabled={node.kind === 'aggregate'} onClick={() => copyPath(node)}><Copy20Regular /><span>Copy path</span><kbd>Alt+Insert</kbd></button>
    {node.kind !== 'aggregate' && <>
      <button onClick={() => { setContext(null); setMenu(null); announce(`${node.name} would be focused in the active Salamander pane.`); }}><Search20Regular /><span>Focus in pane</span></button>
      {node.kind === 'folder' && <button onClick={() => { setContext(null); setMenu(null); announce(`${node.name} would open in the other Salamander pane.`); }}><WindowNew20Regular /><span>Open folder in other pane</span></button>}
      <div className="menu-separator" />
      <button onClick={() => toggleExclusion(node)}>{isExcluded(node, exclusions) ? <Eye20Regular /> : <EyeOff20Regular />}<span>{isExcluded(node, exclusions) ? 'Restore to calculation' : 'Exclude from view'}</span></button>
      <button disabled={isExcluded(node, exclusions) || node.incomplete} onClick={() => { toggleMark(node); setMenu(null); setContext(null); }}><CheckmarkCircle20Regular /><span>{marks.has(node.id) ? 'Remove cleanup mark' : 'Mark for cleanup'}</span></button>
    </>}
  </>;
  const exportRows = rows.map(n => ({ name: n.name, path: n.kind === 'aggregate' ? null : n.path, kind: n.kind, excluded: isExcluded(n, exclusions), summarized: n.kind === 'aggregate',
    sizeBytes: n.incomplete ? null : amount(n, 'size', empty, removed), fileCount: n.incomplete ? null : n.files, sizeOnDiskBytes: unknownDisk(n, empty, removed) ? null : amount(n, 'disk', empty, removed), modified: n.kind === 'aggregate' ? null : n.modified }));
  const exportContent = exportType === 'json' ? JSON.stringify({ proposal: true, path: current.path, metric: metricNames[metric], discoveredBytes: discovered,
    shownBytes, filter: { name: search, type: typeFilter, age: ageFilter, dateBasis }, completeness: scenario === 'empty' ? 'Complete: empty folder' : 'Partial: one unreadable folder; some details summarized', items: exportRows }, null, 2) :
    ['# ViewerSpace proposal; invented data; partial scan', `# Metric: ${metricNames[metric]}; date basis: ${dateBasis}; shown bytes: ${shownBytes}; discovered bytes: ${discovered}`,
      'name,path,kind,excluded,summarized,sizeBytes,fileCount,sizeOnDiskBytes,modified', ...exportRows.map(n => Object.values(n).map(value => value === null ? 'Unavailable' : `"${String(value).replaceAll('"', '""')}"`).join(','))].join('\n');
  const downloadExport = () => {
    const url = URL.createObjectURL(new Blob([exportContent], { type: exportType === 'json' ? 'application/json' : 'text/csv;charset=utf-8' }));
    const link = document.createElement('a'); link.href = url; link.download = `Orion-space-proposal.${exportType}`; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); announce('Export downloaded');
  };

  return <main className={`prototype ${dark ? 'dark' : ''}`}>
    <aside className="proposal-bar"><span className="proposal-label"><span className="proposal-dot" />INTERACTIVE PROPOSAL<span className="proposal-note">Invented data · no live file operations</span></span>
      <div className="prototype-controls"><label>Preview state <select aria-label="Preview state" value={scenario} onChange={event => { setScenario(event.target.value); setClosed(false); if (event.target.value === 'empty') { setSelectedId(null); setExclusions(new Set()); setMarks(new Set()); clearFilters(); } }}>
        <option value="ready">Scan complete</option><option value="scanning">Scanning</option><option value="refreshing">Refreshing</option><option value="network">Slow / partial scan</option><option value="empty">Empty folder</option><option value="partial">Cancelled scan</option>
      </select></label><IconButton icon={dark ? WeatherSunny20Regular : WeatherMoon20Regular} title={dark ? 'Preview light theme' : 'Preview dark theme'} onClick={() => setDark(!dark)} /><button className="text-button reset" onClick={reset}>Reset demo</button></div>
    </aside>
    {closed ? <div className="closed-state"><DataTreemap24Regular /><h1>Viewer closed</h1><p>Escape closes an idle viewer after menus and filters are dismissed.</p><button className="primary" onClick={() => setClosed(false)}>Reopen prototype</button></div> : <section className="app-window" aria-label="Space future workspace">
      <header className="app-title"><DataTreemap24Regular /><strong>{current.name}</strong><span>— Space</span><span className="window-caption">Future workspace</span></header>
      <nav className="menubar" aria-label="Application menu">{['File', 'Navigate', 'View'].map(label => <div className="menu-anchor" key={label}>
        <button aria-expanded={menu === label} onClick={() => setMenu(menu === label ? null : label)}>{label}</button>
        {menu === label && <div className="menu" role="menu">
          {label === 'File' && <><button onClick={() => copyPath()}><Copy20Regular /><span>Copy path</span><kbd>Alt+Insert</kbd></button><button onClick={() => menuAction(() => setModal({ kind: 'export' }))}><ArrowDownload20Regular /><span>Export current view…</span></button><button onClick={() => menuAction(() => setModal({ kind: 'review' }))}><Delete20Regular /><span>Review cleanup marks…</span></button><div className="menu-separator" /><button onClick={() => menuAction(() => setClosed(true))}><Dismiss20Regular /><span>Close viewer</span><kbd>Esc</kbd></button></>}
          {label === 'Navigate' && <><button disabled={!historyIndex} onClick={() => menuAction(() => goHistory(-1))}><ArrowLeft20Regular /><span>Back</span><kbd>Alt+←</kbd></button><button disabled={historyIndex === history.length - 1} onClick={() => menuAction(() => goHistory(1))}><ArrowRight20Regular /><span>Forward</span><kbd>Alt+→</kbd></button><button disabled={!current.parent} onClick={() => navigate(current.parent)}><ArrowUp20Regular /><span>Up one level</span><kbd>Backspace</kbd></button>{itemActions(selected || current)}</>}
          {label === 'View' && <><button onClick={() => menuAction(() => { setScenario('refreshing'); announce('Refreshing this folder; previous results remain visible.'); })}><ArrowClockwise20Regular /><span>Refresh this folder</span></button><button onClick={() => menuAction(() => { setScenario('scanning'); announce('Full scan started in this proposal.'); })}><ArrowClockwise20Regular /><span>Rebuild full scan</span></button><button disabled={!exclusions.size} onClick={() => menuAction(() => setExclusions(new Set()))}><Eye20Regular /><span>Restore all excluded items</span></button><div className="menu-separator" /><button onClick={() => menuAction(() => setModal({ kind: 'settings' }))}><Settings20Regular /><span>View settings…</span></button></>}
        </div>}
      </div>)}</nav>
      <section className="scan-summary" aria-label="Scan summary">
        <div className="primary-summary"><span className="eyebrow">{scenario === 'empty' ? 'DISCOVERED' : scanning || scenario === 'partial' ? 'DISCOVERED SO FAR' : 'DISCOVERED'}</span>
          <strong>{scenario === 'empty' ? '0 B' : formatted(discovered)}</strong><span className="summary-count">{scenario === 'empty' ? '0 files · 0 folders' : `${fixture.files.toLocaleString('en-US')} files · ${folderCount} folders`}</span></div>
        {(exclusions.size > 0 || current.id !== fixture.id) && <div className="shown-summary"><span className="eyebrow">{current.id !== fixture.id ? 'THIS FOLDER · SHOWN' : 'SHOWN'}</span><strong>{formatted(shownBytes)}</strong><span>{exclusions.size ? `${formatted(originalBytes - shownBytes)} excluded in this folder` : 'Included in discovered total'}</span></div>}
        <div className="scan-state"><span className={`status-title ${scanning || scenario === 'refreshing' ? 'working' : ''}`}>{scanning || scenario === 'refreshing' ? <span className="scan-spinner" /> : scenario === 'partial' ? <Warning20Regular /> : <CheckmarkCircle20Regular />}{scanning ? 'Scanning…' : scenario === 'refreshing' ? 'Refreshing this folder…' : scenario === 'partial' ? 'Scan cancelled · partial results' : 'Scan complete'}</span>
          <span>{scanning ? `${scenario === 'network' ? '640' : '12,480'} items/s · exploring folder prioritized` : scenario === 'refreshing' ? 'Showing cached result · 2 min old' : scenario === 'empty' ? 'This folder is empty' : scenario === 'partial' ? 'Unvisited entries may remain · results are partial' : 'Completed just now · partial detail retained'}</span>
          {(scanning || scenario === 'refreshing') && <IconButton icon={Dismiss20Regular} title="Cancel scan (Esc)" onClick={() => { setScenario('partial'); announce('Scan cancelled; current results kept.'); }} />}
        </div>
      </section>
      <div className="workspace-toolbar">
        <div className="navigation-buttons"><IconButton icon={ArrowLeft20Regular} title="Back (Alt+Left)" disabled={!historyIndex} onClick={() => goHistory(-1)} /><IconButton icon={ArrowRight20Regular} title="Forward (Alt+Right)" disabled={historyIndex === history.length - 1} onClick={() => goHistory(1)} /><IconButton icon={ArrowUp20Regular} title="Up one level (Backspace)" disabled={!current.parent} onClick={() => navigate(current.parent)} /></div>
        <nav className="breadcrumbs" aria-label="Folder path"><span className="drive-label">D:</span>{ancestors.map((node, index) => <span key={node.id}><ChevronRight16Regular /><button aria-current={index === ancestors.length - 1 ? 'location' : undefined} onClick={() => index !== ancestors.length - 1 && navigate(node.id)}>{node.name}</button></span>)}</nav>
        <div className="toolbar-actions">
          <label className="metric-control"><span>Area</span><select aria-label="Map area" value={metric} onChange={event => setMetric(event.target.value)}>{Object.entries(metricNames).map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select></label>
          <div className="menu-anchor"><IconButton icon={Color20Regular} title={`Color: ${colorNames[colorMode]}`} aria-expanded={menu === 'colors'} onClick={() => setMenu(menu === 'colors' ? null : 'colors')} />{menu === 'colors' && <div className="menu align-right" role="menu"><p className="menu-caption">Color by</p>{Object.entries(colorNames).map(([key, label]) => <span key={key}>{choice(label, key, colorMode, setColorMode)}</span>)}</div>}</div>
          <div className="menu-anchor"><IconButton icon={layout === 'hidden' ? FullScreenMaximize20Regular : effectiveLayout === 'bottom' ? PanelBottom20Regular : PanelRight20Regular} title={`Layout: ${layout === 'hidden' ? 'map only' : `panel ${effectiveLayout}`}`} aria-expanded={menu === 'layout'} onClick={() => setMenu(menu === 'layout' ? null : 'layout')} />{menu === 'layout' && <div className="menu align-right" role="menu"><p className="menu-caption">Workspace layout</p>{choice('Panel right', 'right', layout, setLayout, PanelRight20Regular)}{choice('Panel bottom', 'bottom', layout, setLayout, PanelBottom20Regular)}{choice('Map only', 'hidden', layout, setLayout, FullScreenMaximize20Regular)}<p className="menu-hint">Ctrl+1 / Ctrl+2 · compact windows use bottom</p></div>}</div>
        </div>
      </div>
      <div className="view-context"><div><span className="context-label">{metricNames[metric]} treemap</span><span className="dot-divider">·</span><span>{metric === 'disk' ? 'Known allocation only; unknown items stay in the list' : 'Click a folder to explore'}</span>{activeFilters && <span className="filter-badge">Filter highlights · totals unchanged</span>}</div>
        {exclusions.size > 0 && <div className="exclusion-summary"><EyeOff20Regular /><span>{exclusions.size} excluded</span><button onClick={() => setExclusions(new Set())}>Restore all</button></div>}
      </div>
      {metric === 'disk' && <div className="allocation-notice"><Info20Regular /><span>Allocation is incomplete. Tile area and shares use known size on disk; unknown allocation is never replaced with Size.</span></div>}
      <div className={`workspace layout-${effectiveLayout}`}>
        <section className="map-pane" aria-label="Treemap exploration"><div className="map-canvas">
          {scenario === 'empty' ? <div className="map-empty"><Folder20Regular /><strong>This folder is empty</strong><span>Choose another folder to explore.</span></div> : <MapView current={current} metric={metric} colorMode={colorMode} exclusions={exclusions} removed={removed} selected={selectedId} marks={marks} matches={matches} onSelect={setSelectedId} onEnter={navigate} onContext={contextFor} />}
        </div><div className="map-legend"><span className="legend-label">{colorNames[colorMode]}</span>{colorMode === 'type' ? Object.entries(categories).map(([key, cat]) => <button key={key} className={typeFilter === key ? 'active' : ''} onClick={() => setTypeFilter(typeFilter === key ? '' : key)}><i style={{ background: cat.color }} />{cat.label === 'Unclassified / summarized' ? 'Unclassified' : cat.label}</button>) : colorMode === 'age' || colorMode === 'newest' ? <>{['Last 7 days', '7–30 days', '30 days–1 year', 'Over 1 year'].map((label, i) => <span key={label}><i style={{ background: ageColors[i] }} />{label}</span>)}<span className="legend-disclosure">{colorMode === 'newest' ? 'Newest known file inside; summarized dates may be unavailable' : 'Folder’s own Modified date'}</span></> : <span>Each top-level group has its own color<span className="legend-short"> · Other files are summarized</span></span>}</div></section>
        {layout !== 'hidden' && <aside className="analysis-panel" aria-label="Folder analysis and selection"><div className="analysis-main">
          <div className="panel-heading"><h2>In this folder</h2><span>{children.length} entries</span><div className="menu-anchor"><IconButton icon={MoreHorizontal20Regular} title="Folder actions" aria-expanded={menu === 'folder-actions'} onClick={() => setMenu(menu === 'folder-actions' ? null : 'folder-actions')} />{menu === 'folder-actions' && <div className="menu align-right" role="menu"><button onClick={markMatching}><CheckmarkCircle20Regular /><span>Mark matching items</span></button><button onClick={() => menuAction(() => setModal({ kind: 'export' }))}><ArrowDownload20Regular /><span>Export current view…</span></button><button disabled={!exclusions.size} onClick={() => menuAction(() => setExclusions(new Set()))}><Eye20Regular /><span>Restore all excluded items</span></button></div>}</div></div>
          <div className="filter-row"><div className="search-box"><Search20Regular /><input ref={filterRef} aria-label="Filter this folder by name" placeholder="Filter this folder by name" value={search} onChange={event => setSearch(event.target.value)} />{search && <IconButton icon={Dismiss20Regular} title="Clear name filter" onClick={() => setSearch('')} />}</div><div className="menu-anchor"><IconButton icon={Filter20Regular} title="Type and age filters" aria-expanded={menu === 'filters'} className={`icon-button filter-toggle ${typeFilter || ageFilter ? 'active' : ''}`} onClick={() => setMenu(menu === 'filters' ? null : 'filters')} />{menu === 'filters' && <div className="menu filter-menu align-right" role="group" aria-label="Current-folder filters"><h3>Filter this folder</h3><label>File type<select value={typeFilter} onChange={event => setTypeFilter(event.target.value)} aria-label="File type filter"><option value="">All file types</option>{Object.entries(categories).map(([key, cat]) => <option key={key} value={key}>{cat.label}</option>)}</select></label><label>Age<select value={ageFilter} onChange={event => setAgeFilter(event.target.value)} aria-label="Age filter"><option value="">Any age</option><option value="month">Not modified in 30 days</option><option value="year">Not modified in one year</option></select></label><label>Date basis<select value={dateBasis} onChange={event => setDateBasis(event.target.value)} aria-label="Age filter date basis"><option value="modified">Item’s Modified date</option><option value="newest">Newest known file inside folders</option></select></label><p>As of 22 Oct 2026. Folders match when retained contents contain the chosen type. Unknown dates do not match an age filter.</p><button onClick={() => { clearFilters(); setMenu(null); }}>Clear filters</button><button className="primary" onClick={() => setMenu(null)}>Apply</button></div>}</div></div>
          <p className="scope-note" title="Search uses retained names in this folder. Names summarized as Other files are unavailable. Filtering does not change map sizes or percentages.">Retained names only · summarized items are not searchable</p>
          {(typeFilter || ageFilter) && <div className="filter-chips">{typeFilter && <button onClick={() => setTypeFilter('')}>{categories[typeFilter].label}<Dismiss20Regular /></button>}{ageFilter && <button onClick={() => setAgeFilter('')}>{ageFilter === 'year' ? 'Over 1 year' : 'Over 30 days'} · {dateBasis === 'modified' ? 'Modified' : 'Newest inside'}<Dismiss20Regular /></button>}</div>}
          <button className="breakdown-toggle" aria-expanded={showTypes} onClick={() => setShowTypes(!showTypes)}><ChevronRight16Regular className={showTypes ? 'expanded' : ''} /><span>File types</span><span className="muted">{types.length} groups</span></button>
          {showTypes && <div className="type-breakdown">{types.map(([key, value]) => <button key={key} className={typeFilter === key ? 'active' : ''} onClick={() => setTypeFilter(typeFilter === key ? '' : key)}><i style={{ background: categories[key].color }} /><span>{categories[key].label}<small>{value.files.toLocaleString('en-US')} files</small></span><span>{formatted(value.bytes)}</span><span className="type-bar" style={{ width: `${shownBytes ? value.bytes / shownBytes * 100 : 0}%`, background: categories[key].color }} /></button>)}</div>}
          <div className="item-list" role="region" aria-label="Ranked items"><table><thead><tr><th className="mark-cell"><span className="sr-only">Cleanup mark</span></th><th><button onClick={() => setSortBy('name')}>Name {sortBy === 'name' && <ChevronDown16Regular />}</button></th><th className="number"><button onClick={() => setSortBy('size')}>{metricNames[metric]} {sortBy === 'size' && <ChevronDown16Regular />}</button></th><th className="number share-col">Share</th></tr></thead><tbody>{rows.map(node => {
            const excluded = isExcluded(node, exclusions), unknown = node.incomplete || (metric === 'disk' && unknownDisk(node, empty, removed));
            const value = amount(node, metric, excluded ? empty : exclusions, removed);
            return <tr key={node.id} className={`${selectedId === node.id ? 'selected' : ''} ${excluded ? 'excluded' : ''}`} aria-selected={selectedId === node.id} onClick={() => setSelectedId(node.id)} onDoubleClick={() => node.kind === 'folder' && navigate(node.id)} onContextMenu={event => { event.preventDefault(); contextFor(node, event.clientX, event.clientY); }} tabIndex={0} onKeyDown={event => { if (event.key === 'Enter' && event.target === event.currentTarget && node.kind === 'folder') navigate(node.id); if (event.key === 'ContextMenu' || (event.shiftKey && event.key === 'F10')) { event.preventDefault(); const r = event.currentTarget.getBoundingClientRect(); contextFor(node, r.x + 30, r.y + 20); } }}>
              <td className="mark-cell"><input type="checkbox" checked={marks.has(node.id) || [...marks].some(id => isInside(node.id, id))} disabled={excluded || node.kind === 'aggregate' || node.incomplete || [...marks].some(id => id !== node.id && isInside(node.id, id))} aria-label={`Mark ${node.name} for cleanup`} onClick={event => event.stopPropagation()} onChange={() => toggleMark(node)} /></td>
              <td className="name-cell"><div>{node.kind !== 'aggregate' ? <button className="visibility-target" title={`${excluded ? 'Restore' : 'Exclude'} ${node.name} ${excluded ? 'to calculation' : 'from view'}`} aria-label={`${excluded ? 'Restore' : 'Exclude'} ${node.name}`} onClick={event => { event.stopPropagation(); toggleExclusion(node); }}><ItemIcon node={node} excluded={excluded} /></button> : <ItemIcon node={node} />}<span className="item-name" title={node.path}>{node.name}</span>{node.issue && <Warning20Regular className="warning-icon" aria-label="Unreadable folder" />}{excluded && <span className="row-state">Excluded</span>}</div></td>
              <td className="number" title={unknown ? 'Allocation or folder size unavailable' : `${value.toLocaleString('en-US')} ${metric === 'count' ? 'files' : 'bytes'}`}>{node.incomplete ? 'Unavailable' : `${unknown && value ? '≥ ' : ''}${unknown && !value ? 'Unavailable' : metricText(value, metric)}`}</td><td className="number share-col">{excluded || unknown && !value ? '—' : pct(value, visible)}</td>
            </tr>;
          })}</tbody></table>{!rows.length && <div className="list-empty"><Search20Regular /><strong>{activeFilters ? 'No retained items match' : 'No items in this folder'}</strong>{activeFilters && <button className="text-button" onClick={clearFilters}>Clear filters</button>}</div>}</div>
        </div>
        <section className="inspector" aria-label="Selection inspector"><div className="inspector-heading"><h2 className={selected ? 'sr-only' : ''}>Selection</h2>{selected && <><div className="selection-name"><ItemIcon node={selected} excluded={selectionExcluded} /><strong title={selected.name}>{selected.name}</strong></div><IconButton icon={Copy20Regular} title="Copy selected path (Alt+Insert)" disabled={selected.kind === 'aggregate'} onClick={() => copyPath(selected)} /><div className="menu-anchor"><IconButton icon={MoreHorizontal20Regular} title="Selection actions" aria-expanded={menu === 'selection'} onClick={() => setMenu(menu === 'selection' ? null : 'selection')} />{menu === 'selection' && <div className="menu align-right upwards" role="menu">{itemActions(selected)}</div>}</div></>}</div>
          {selected ? <><p className="selection-location" title={selected.path}>{parentName(selected)}</p><div className="selection-metrics"><div><span>{metricNames[metric]}</span><strong title={`${selectionValue.toLocaleString('en-US')} ${metric === 'count' ? 'files' : 'bytes'}`}>{selected.incomplete ? 'Unavailable' : selectionUnknown && !selectionValue ? 'Unavailable' : `${selectionUnknown ? '≥ ' : ''}${metricText(selectionValue, metric)}`}</strong></div><div><span>{selectionExcluded ? 'Calculation' : '% of folder'}</span><strong>{selected.incomplete ? 'Unavailable' : selectionExcluded ? 'Excluded' : pct(selectionValue, visible)}</strong></div><div className="modified"><span>Modified</span><strong>{selected.kind === 'aggregate' ? 'Unavailable' : selected.modified}</strong></div></div>
            {(selectionExcluded || selected.incomplete || selected.kind === 'aggregate' || scenario === 'refreshing') && <p className="selection-state"><Info20Regular />{selectionExcluded ? 'Excluded from shown sizes' : selected.incomplete ? 'Unreadable. This folder’s size is unknown.' : selected.kind === 'aggregate' ? 'Exact total; individual names were summarized.' : 'Cached result · refresh in progress'}</p>}
            <button className="details-toggle" aria-expanded={details} onClick={() => setDetails(!details)}>{details ? 'Fewer details' : 'More details'}<ChevronDown16Regular className={details ? 'expanded' : ''} /></button>{details && <dl className="extra-details"><dt>Size on disk</dt><dd>{unknownDisk(selected, empty, removed) ? 'Unavailable (incomplete coverage)' : formatted(amount(selected, 'disk', empty, removed))}</dd><dt>Files</dt><dd>{selected.incomplete ? 'Unavailable' : selected.files.toLocaleString('en-US')}</dd><dt>Newest known file</dt><dd>{selected.newest || 'Unavailable'}{selected.children && unknownDisk(selected, empty, removed) ? ' · partial detail' : ''}</dd><dt>Exact size</dt><dd>{selected.incomplete ? 'Unavailable' : `${selected.bytes.toLocaleString('en-US')} bytes`}</dd></dl>}
          </> : <p className="selection-empty">Select an item to inspect its size and dates.<br />Double-click a folder in the list to explore it.</p>}
        </section></aside>}
      </div>
      <footer className="statusbar"><button className="problem-button" disabled={scenario === 'empty'} onClick={() => setModal({ kind: 'problems' })}><Warning20Regular /><span>{scenario === 'empty' ? 'No scan problems' : '1 unreadable folder'}</span><ChevronRight16Regular /></button><span className="retention-state" hidden={scenario === 'empty'} title="All measured totals are retained. Some individual names and allocation details have been summarized within the detail budget.">Some names summarized · totals retained</span><span className="status-spacer" /><button className="export-button" title="Export current view" onClick={() => setModal({ kind: 'export' })}><ArrowDownload20Regular /><span>Export</span></button><button className={`review-button ${marks.size ? 'has-marks' : ''}`} onClick={() => setModal({ kind: 'review' })}><CheckmarkCircle20Regular /><span>Review cleanup{marks.size ? ` (${marks.size})` : ''}</span></button></footer>
    </section>}
    {(menu || context) && <div className="menu-dismiss" onClick={() => { setMenu(null); setContext(null); }} />}
    {context && <div className="menu context-menu" style={{ left: context.x, top: context.y }} role="menu" aria-label={`Actions for ${context.node.name}`}>{itemActions(context.node)}</div>}
    {toast && <div className="toast" role="status"><CheckmarkCircle20Regular />{toast}<IconButton icon={Dismiss20Regular} title="Dismiss notification" onClick={() => setToast('')} /></div>}
    {modal && <Modal title={{ review: 'Review cleanup', export: 'Export current view', problems: 'Scan problems', copy: 'Copy path', settings: 'View settings' }[modal.kind]} onClose={closeModal} className={modal.kind === 'export' ? 'export-dialog' : ''}>
      {modal.kind === 'review' && <><div className="dialog-body"><p className="dialog-intro">Review the items you marked before handing them to File Operations.</p>{markedNodes.length ? <><div className="review-summary"><strong>{markedNodes.length} {markedNodes.length === 1 ? 'target' : 'targets'} · {formatted(markedBytes)}</strong><span>Folder targets include all contents, including files outside the filter. Nested targets count once.</span></div><div className="review-items">{markedNodes.map(n => <div key={n.id}><ItemIcon node={n} /><span><strong>{n.name}</strong><small>{parentName(n)}</small></span><span>{formatted(amount(n, 'size', empty, removed))}</span><IconButton icon={Dismiss20Regular} title={`Remove ${n.name} from cleanup`} onClick={() => toggleMark(n)} /></div>)}</div><fieldset className="operation-choice"><legend>Requested operation</legend><label><input type="radio" name="operation" value="recycle" checked={deleteMode === 'recycle'} onChange={() => setDeleteMode('recycle')} />Move to Recycle Bin</label><label><input type="radio" name="operation" value="delete" checked={deleteMode === 'delete'} onChange={() => setDeleteMode('delete')} />Delete permanently</label></fieldset></> : <div className="review-empty"><CheckmarkCircle20Regular /><strong>No items marked</strong><p>Use the checkboxes in the list, or filter the current folder and choose “Mark matching items” from its menu.</p></div>}<p className="prototype-disclosure"><Info20Regular />This proposal simulates the handoff only. It cannot delete or recycle real files.</p></div><footer><button onClick={closeModal}>Keep exploring</button><button className="primary" disabled={!marks.size} onClick={() => { setMarks(new Set()); closeModal(); announce(`${markedNodes.length} ${markedNodes.length === 1 ? 'target' : 'targets'} handed to File Operations (simulation). No real files changed.`); }}>Send to File Operations</button></footer></>}
      {modal.kind === 'export' && <><div className="dialog-body"><p className="dialog-intro">Export the current folder’s retained entries with exclusion and completeness information.</p><div className="export-options"><label>Format<select aria-label="Export format" value={exportType} onChange={event => setExportType(event.target.value)}><option value="json">JSON</option><option value="csv">CSV</option></select></label><span>{rows.length} entries · {current.path}</span></div><pre className="export-preview" tabIndex={0}>{exportContent}</pre><p className="muted">Includes active filters, excluded rows and unavailable allocation. The browser will save a file with invented data.</p></div><footer><button onClick={closeModal}>Cancel</button><button className="primary" onClick={downloadExport}><ArrowDownload20Regular />Download {exportType.toUpperCase()}</button></footer></>}
      {modal.kind === 'problems' && <><div className="dialog-body"><p className="dialog-intro">The result is usable, but it is not a complete account of every file.</p><div className="problem-entry"><Warning20Regular /><div><strong>{problemNode.name}</strong><p>{problemNode.path}</p><span>Access denied · contents and size unavailable</span></div><button onClick={() => { navigate(fixture.id); setSelectedId(problemNode.id); closeModal(); }}>Locate</button></div><div className="problem-info"><Info20Regular /><div><strong>Some names were summarized</strong><p>Measured sizes and counts are retained under Other files. Their individual names, dates and allocation details are unavailable.</p></div></div><p>Unreadable folders remain in the list. Administrator rescanning is outside ViewerSpace.</p></div><footer><button className="primary" onClick={closeModal}>Close</button></footer></>}
      {modal.kind === 'copy' && <><div className="dialog-body"><p>Clipboard access is unavailable in this preview. Select and copy the full path:</p><input className="copy-field" aria-label="Full path" value={modal.node.path} readOnly onFocus={event => event.target.select()} /></div><footer><button className="primary" onClick={closeModal}>Done</button></footer></>}
      {modal.kind === 'settings' && <><div className="dialog-body"><p className="dialog-intro">Keep everyday display choices close. Technical scan options stay under Advanced.</p><label className="setting-row">Default area<select aria-label="Default area" value={metric} onChange={event => setMetric(event.target.value)}>{Object.entries(metricNames).map(([key, label]) => <option value={key} key={key}>{label}</option>)}</select></label><label className="setting-row">Color by<select aria-label="Default colors" value={colorMode} onChange={event => setColorMode(event.target.value)}>{Object.entries(colorNames).map(([key, label]) => <option value={key} key={key}>{label}</option>)}</select></label><button className="breakdown-toggle" aria-expanded={advanced} onClick={() => setAdvanced(!advanced)}><ChevronRight16Regular className={advanced ? 'expanded' : ''} />Advanced</button>{advanced && <div className="advanced-options"><dl><dt>Scan workers</dt><dd>Configured count · unchanged</dd><dt>Memory budget</dt><dd>Stop/refuse if the configured budget cannot fit</dd><dt>Retention</dt><dd>Largest entries first · exact totals preserved</dd><dt>Scan reuse</dt><dd>In memory for this viewer only</dd><dt>Minimized scan</dt><dd>Lower background priority · same worker count</dd></dl><p>Prototype values are descriptive. This view does not edit application settings.</p></div>}</div><footer><button className="primary" onClick={closeModal}>Done</button></footer></>}
    </Modal>}
  </main>;
}
