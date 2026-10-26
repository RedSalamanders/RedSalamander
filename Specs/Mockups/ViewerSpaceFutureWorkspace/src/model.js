export const GB = 1024 ** 3;
let sequence = 0;
const file = (name, gb, type, modified = '2026-10-18 14:32', options = {}) => ({
  id: `item-${++sequence}`, name, kind: 'file', bytes: Math.round(gb * GB),
  disk: Math.round(gb * GB), files: 1, type, modified, ...options,
});
const other = (gb, files) => ({ ...file('Other files', gb, 'other'), kind: 'aggregate', files, disk: null, modified: '' });
const folder = (name, children, modified = '2026-10-21 10:24', options = {}) => ({
  id: `item-${++sequence}`, name, kind: 'folder', children, modified, ...options,
});
export const fixture = folder('Orion', [
  folder('build', [
    folder('release', [file('orion-symbols.pdb', 7.8, 'build'), file('orion-app.msix', 4.2, 'archives'), other(3.1, 18240)]),
    folder('debug', [file('orion-debug.pdb', 5.6, 'build'), file('engine.pch', 3.4, 'build'), other(1.8, 24360)]),
    file('compile-trace.json', .34, 'data'), other(1.46, 5280),
  ]),
  folder('.cache', [
    folder('packages', [file('sdk-bundle.7z', 6.1, 'archives'), other(4.3, 31520)]),
    folder('thumbnails', [other(3.7, 18420)]),
    file('shader-cache.bin', 2.6, 'build', '2026-08-12 09:10'), other(1.3, 8400),
  ], '2026-10-22 08:52'),
  folder('assets', [
    folder('video', [file('launch-film.mov', 6.4, 'media', '2026-10-02 16:42'), file('product-walkthrough.mp4', 2.8, 'media')]),
    folder('images', [file('environment-textures.exr', 2.7, 'media'), file('brand-photos.zip', 1.4, 'archives'), other(.85, 1836)]),
    file('illustrations.fig', .55, 'documents'), other(.3, 418),
  ], '2026-10-15 15:18'),
  folder('packages', [file('runtime-sdk.7z', 4.1, 'archives'), file('toolchain.zip', 2.7, 'archives'), other(1.2, 6480)]),
  folder('src', [
    folder('engine', [file('renderer.cpp', .00012, 'code'), file('scene-index.bin', 1.1, 'data'), other(.28, 5240)]),
    folder('workspace', [file('App.tsx', .00004, 'code'), other(.62, 1380)]),
    file('project-index.db', .8, 'data'), other(.2, 980),
  ]),
  folder('.git', [file('objects.pack', 2.9, 'data'), other(.1, 1630)]),
  folder('docs', [file('architecture.pdf', .12, 'documents'), file('design-handbook.pdf', .08, 'documents'), other(.06, 84)]),
  file('release-backup.7z', 7.6, 'archives', '2025-09-08 12:10'),
  file('sample-dataset.parquet', 2.8, 'data', '2026-09-18 11:08', { disk: Math.round(1.9 * GB) }),
  file('walkthrough.mp4', 2.4, 'media', '2026-07-16 09:40'),
  file('README.md', .00002, 'code'),
  folder('restricted-data', [], '2026-08-02 11:24', { issue: 'Access denied', incomplete: true }),
], '2026-10-22 09:12');

export const categories = {
  build: { label: 'Build files', color: '#77a8db' },
  archives: { label: 'Archives', color: '#b49bd2' },
  media: { label: 'Media', color: '#e4aa79' },
  data: { label: 'Data', color: '#6eb7b1' },
  code: { label: 'Source code', color: '#97ba7c' },
  documents: { label: 'Documents', color: '#d9bc68' },
  other: { label: 'Unclassified / summarized', color: '#bdc5c9' },
};
export const groupColors = ['#83afa3', '#8fabce', '#b8a0cd', '#d6b38a', '#9db98e', '#93bac9', '#ccb793', '#a9acd2', '#97c2b8', '#dcbaa1'];
export const nodes = new Map();
function attach(node, parent = null) {
  node.parent = parent?.id ?? null;
  node.path = parent ? `${parent.path}\\${node.name}` : 'D:\\Workspaces\\Orion';
  if (node.children) {
    node.children.forEach(child => attach(child, node));
    node.bytes = node.children.reduce((sum, child) => sum + child.bytes, 0);
    node.files = node.children.reduce((sum, child) => sum + child.files, 0);
  }
  node.newest = (node.children?.reduce((latest, child) => child.newest > latest ? child.newest : latest, '') || (node.kind === 'file' ? node.modified : ''));
  nodes.set(node.id, node);
}
attach(fixture);

export function isInside(id, parentId) {
  let current = nodes.get(id);
  while (current) { if (current.id === parentId) return true; current = nodes.get(current.parent); }
  return false;
}
export function isExcluded(node, exclusions) {
  return [...exclusions].some(id => isInside(node.id, id));
}
export function amount(node, metric = 'size', exclusions = new Set(), removed = new Set()) {
  if (removed.has(node.id) || isExcluded(node, exclusions)) return 0;
  if (node.children) return node.children.reduce((sum, child) => sum + amount(child, metric, exclusions, removed), 0);
  return metric === 'count' ? node.files : metric === 'disk' ? (node.disk ?? 0) : node.bytes;
}
export function unknownDisk(node, exclusions = new Set(), removed = new Set()) {
  if (removed.has(node.id) || isExcluded(node, exclusions)) return false;
  return node.incomplete || (node.children ? node.children.some(child => unknownDisk(child, exclusions, removed)) : node.disk === null);
}
export function formatted(bytes) {
  if (bytes >= GB) return `${(bytes / GB).toLocaleString('en-US', { maximumFractionDigits: 1 })} GB`;
  if (bytes >= 1024 ** 2) return `${(bytes / 1024 ** 2).toLocaleString('en-US', { maximumFractionDigits: 1 })} MB`;
  if (bytes >= 1024) return `${Math.round(bytes / 1024).toLocaleString('en-US')} KB`;
  return `${bytes.toLocaleString('en-US')} B`;
}
export const metricText = (value, metric) => metric === 'count' ? `${value.toLocaleString('en-US')} ${value === 1 ? 'file' : 'files'}` : formatted(value);
export function breakdown(node, exclusions, removed) {
  const result = Object.fromEntries(Object.keys(categories).map(type => [type, { bytes: 0, files: 0 }]));
  const walk = item => {
    if (removed.has(item.id) || isExcluded(item, exclusions)) return;
    if (item.children) item.children.forEach(walk);
    else { result[item.type].bytes += item.bytes; result[item.type].files += item.files; }
  };
  walk(node);
  return Object.entries(result).filter(([, item]) => item.files || item.bytes).sort((a, b) => b[1].bytes - a[1].bytes);
}
export function hasType(node, type) {
  return !type || (node.children ? node.children.some(child => hasType(child, type)) : node.type === type);
}

// Squarified data geometry: rectangles represent the chosen measured metric.
export function treemap(items, width, height, getValue) {
  const weighted = items.map(node => ({ node, value: getValue(node) })).filter(item => item.value > 0).sort((a, b) => b.value - a.value);
  const total = weighted.reduce((sum, item) => sum + item.value, 0);
  if (!total || width <= 0 || height <= 0) return [];
  const pending = weighted.map(item => ({ ...item, area: item.value / total * width * height }));
  let x = 0, y = 0, w = width, h = height, row = [], result = [];
  const worst = (group, side) => {
    if (!group.length || side <= 0) return Infinity;
    const sum = group.reduce((s, item) => s + item.area, 0);
    const max = Math.max(...group.map(item => item.area)), min = Math.min(...group.map(item => item.area));
    return Math.max(side * side * max / (sum * sum), sum * sum / (side * side * min));
  };
  const place = () => {
    const area = row.reduce((sum, item) => sum + item.area, 0);
    if (w >= h) {
      const strip = area / h; let dy = y;
      row.forEach(item => { const rh = item.area / strip; result.push({ node: item.node, x, y: dy, w: strip, h: rh }); dy += rh; });
      x += strip; w = Math.max(0, w - strip);
    } else {
      const strip = area / w; let dx = x;
      row.forEach(item => { const rw = item.area / strip; result.push({ node: item.node, x: dx, y, w: rw, h: strip }); dx += rw; });
      y += strip; h = Math.max(0, h - strip);
    }
    row = [];
  };
  while (pending.length) {
    const side = Math.min(w, h), next = pending[0];
    if (!row.length || worst([...row, next], side) <= worst(row, side)) row.push(pending.shift());
    else place();
  }
  if (row.length) place();
  return result;
}
