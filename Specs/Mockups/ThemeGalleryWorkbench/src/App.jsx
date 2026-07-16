import { useMemo, useState } from "react";
import {
  Archive,
  ArrowRight,
  Braces,
  Check,
  ChevronDown,
  ChevronRight,
  Copy,
  FileCode,
  FileDown,
  FileSpreadsheet,
  FileText,
  Folder,
  FolderPlus,
  FunctionSquare,
  HardDrive,
  Image,
  Menu,
  Minus,
  Rainbow,
  Scale,
  Search,
  Sparkles,
  Square,
  X,
  Zap,
} from "lucide-react";

const themes = [
  { id: "builtin-light", name: "Built-in Light", family: "Built-in", base: "Light", palette: ["#ffffff", "#f1f1f1", "#d5d5d5", "#0078d4", "#1f1f1f"], colors: { bg: "#ffffff", surface: "#f3f3f3", panel: "#f8f8f8", border: "#d4d4d4", text: "#191919", muted: "#747474", accent: "#0078d4", select: "#cce8ff", selectBorder: "#0078d4" } },
  { id: "builtin-dark", name: "Built-in Dark", family: "Built-in", base: "Dark", dark: true, palette: ["#1e1e1e", "#252526", "#3f3f46", "#4daafc", "#f1f1f1"], colors: { bg: "#1e1e1e", surface: "#252526", panel: "#2d2d30", border: "#3f3f46", text: "#f1f1f1", muted: "#a7a7a7", accent: "#4daafc", select: "#174b71", selectBorder: "#4daafc" } },
  { id: "rainbow-light", name: "Rainbow Light", family: "Built-in dynamic", base: "Light", rainbow: true, palette: ["#ffffff", "#23d795", "#5b9cf6", "#f45bb0", "#ffd166"], colors: { bg: "#ffffff", surface: "#f4f4f4", panel: "#fafafa", border: "#d6d6d6", text: "#161616", muted: "#777777", accent: "#22d893", select: "#d8edff", selectBorder: "#148bd1" } },
  { id: "rainbow-dark", name: "Rainbow Dark", family: "Built-in dynamic", base: "Dark", dark: true, rainbow: true, palette: ["#17181c", "#37e6a5", "#6fb5ff", "#ff70bd", "#f7cf66"], colors: { bg: "#17181c", surface: "#202126", panel: "#25272d", border: "#3d4048", text: "#f6f6f7", muted: "#a7aab2", accent: "#37e6a5", select: "#214e48", selectBorder: "#37e6a5" } },
  { id: "forest-mist", name: "Forest Mist", family: "Existing", base: "Dark", dark: true, palette: ["#101a16", "#183127", "#265e45", "#61d095", "#e3f5eb"], colors: { bg: "#101a16", surface: "#14241d", panel: "#193027", border: "#2b4a3b", text: "#e3f5eb", muted: "#9dbbab", accent: "#61d095", select: "#23533d", selectBorder: "#72dca5" } },
  { id: "neon-tokyo", name: "Neon Tokyo", family: "Existing", base: "Dark", dark: true, palette: ["#0b1020", "#15213c", "#27d9ff", "#ff3fd1", "#f7f6ff"], colors: { bg: "#0b1020", surface: "#101a31", panel: "#14213d", border: "#2d3b62", text: "#f7f6ff", muted: "#96a7cc", accent: "#ff3fd1", select: "#502052", selectBorder: "#ff59dc" } },
  { id: "paper-ink", name: "Paper & Ink", family: "Existing", base: "Light", palette: ["#f7f0df", "#ebe1ca", "#c8b997", "#b34d35", "#2d2922"], colors: { bg: "#f7f0df", surface: "#efe6d3", panel: "#f3ead8", border: "#c8b997", text: "#2d2922", muted: "#766d5e", accent: "#b34d35", select: "#ead0b5", selectBorder: "#b34d35" } },
  { id: "retro-terminal", name: "Retro Terminal", family: "Existing", base: "Dark", dark: true, mono: true, palette: ["#071108", "#0d1e10", "#1e5a29", "#52f36a", "#c6ffd0"], colors: { bg: "#071108", surface: "#0a170c", panel: "#0d1e10", border: "#1e5a29", text: "#c6ffd0", muted: "#68a874", accent: "#52f36a", select: "#153a1c", selectBorder: "#52f36a" } },
  { id: "solar-flare", name: "Solar Flare", family: "Existing", base: "Dark", dark: true, palette: ["#1b1110", "#2c1914", "#5c2d1f", "#ff8a2c", "#fff1df"], colors: { bg: "#1b1110", surface: "#241511", panel: "#2c1914", border: "#5c392b", text: "#fff1df", muted: "#c9a88b", accent: "#ff8a2c", select: "#613216", selectBorder: "#ff9e50" } },
  { id: "ugly", name: "Ugly", family: "Existing", base: "Light", palette: ["#ffff00", "#00ffff", "#ff00ff", "#00ff00", "#000080"], colors: { bg: "#ffff00", surface: "#00ffff", panel: "#ffb6ff", border: "#000080", text: "#000080", muted: "#641064", accent: "#ff00ff", select: "#00ff00", selectBorder: "#000080" } },
  { id: "dracula", name: "Dracula", family: "New · MIT", base: "Dark", dark: true, palette: ["#282a36", "#44475a", "#6272a4", "#bd93f9", "#f8f8f2"], seedChoices: ["#ff5555", "#ffb86c", "#f1fa8c", "#50fa7b", "#8be9fd", "#bd93f9"], colors: { bg: "#282a36", surface: "#30323f", panel: "#343746", border: "#44475a", text: "#f8f8f2", muted: "#a6a8ba", accent: "#bd93f9", link: "#8be9fd", select: "#44475a", selectBorder: "#bd93f9" } },
  { id: "cat-latte", name: "Catppuccin Latte", family: "New · MIT · Lavender", base: "Light", palette: ["#eff1f5", "#dce0e8", "#7c7f93", "#7287fd", "#1e66f5", "#4c4f69"], seedChoices: ["#d20f39", "#fe640b", "#df8e1d", "#40a02b", "#209fb5", "#7287fd"], catppuccin: { accent: "Lavender", action: "Blue", selection: "Overlay2 · 26%" }, colors: { bg: "#eff1f5", surface: "#e6e9ef", panel: "#dce0e8", border: "#acb0be", text: "#4c4f69", muted: "#6c6f85", accent: "#7287fd", link: "#1e66f5", select: "#d1d3dc", selectBorder: "#7287fd", mauve: "#8839ef" } },
  { id: "cat-frappe", name: "Catppuccin Frappé", family: "New · MIT · Lavender", base: "Dark", dark: true, palette: ["#303446", "#232634", "#949cbb", "#babbf1", "#8caaee", "#c6d0f5"], seedChoices: ["#e78284", "#ef9f76", "#e5c890", "#a6d189", "#85c1dc", "#babbf1"], catppuccin: { accent: "Lavender", action: "Blue", selection: "Overlay2 · 26%" }, colors: { bg: "#303446", surface: "#292c3c", panel: "#232634", border: "#626880", text: "#c6d0f5", muted: "#a5adce", accent: "#babbf1", link: "#8caaee", select: "#4a4f64", selectBorder: "#babbf1", mauve: "#ca9ee6" } },
  { id: "cat-mocha", name: "Catppuccin Mocha", family: "New · MIT · Lavender", base: "Dark", dark: true, palette: ["#1e1e2e", "#11111b", "#9399b2", "#b4befe", "#89b4fa", "#cdd6f4"], seedChoices: ["#f38ba8", "#fab387", "#f9e2af", "#a6e3a1", "#74c7ec", "#b4befe"], catppuccin: { accent: "Lavender", action: "Blue", selection: "Overlay2 · 26%" }, colors: { bg: "#1e1e2e", surface: "#181825", panel: "#11111b", border: "#585b70", text: "#cdd6f4", muted: "#a6adc8", accent: "#b4befe", link: "#89b4fa", select: "#3c3e50", selectBorder: "#b4befe", mauve: "#cba6f7" } },
];

const functionOptions = [
  ["seeded-choice", "seededChoice"],
  ["system-color", "systemColor"],
  ["perceptual-tone", "perceptualTone"],
  ["ensure-contrast", "ensureContrast"],
  ["harmonize", "harmonize"],
];

const systemColors = {
  accent: "#0078d4",
  accentLight: "#60cdff",
  accentDark: "#005a9e",
  window: "#ffffff",
  windowText: "#111111",
  highlight: "#0067c0",
  highlightText: "#ffffff",
};

function hexRgb(value) {
  const match = /^#([0-9a-f]{6})$/i.exec(value ?? "");
  if (!match) return null;
  const packed = Number.parseInt(match[1], 16);
  return [(packed >> 16) & 255, (packed >> 8) & 255, packed & 255];
}

function relativeLuminance(value) {
  const rgb = hexRgb(value);
  if (!rgb) return 0;
  const linear = rgb.map((channel) => {
    const unit = channel / 255;
    return unit <= 0.04045 ? unit / 12.92 : ((unit + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
}

function contrastRatio(first, second) {
  const a = relativeLuminance(first);
  const b = relativeLuminance(second);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

function buildFunctionDemo(theme, functionId, { systemRole, toneValue, contrastTarget, harmonizeAmount }) {
  const colors = { ...theme.colors, link: theme.colors.link ?? theme.colors.accent, selectionText: theme.colors.text };
  const demo = { colors, selectionChoices: null, before: colors.accent, after: colors.accent, appliedTo: "Active pane accent", expression: "systemColor(accent)", phase: "event" };

  if (functionId === "seeded-choice") {
    demo.selectionChoices = theme.seedChoices ?? theme.palette.slice(1);
    demo.before = colors.select;
    demo.after = demo.selectionChoices[0];
    demo.appliedTo = "Multiple-selection rows";
    demo.expression = "seededChoice(runtime.seed, palette.red, …)";
    demo.phase = "paint";
  } else if (functionId === "system-color") {
    const selected = systemColors[systemRole];
    colors.accent = selected;
    colors.link = selected;
    colors.selectBorder = selected;
    demo.after = selected;
    demo.expression = `systemColor(${systemRole})`;
  } else if (functionId === "perceptual-tone") {
    const target = Number(toneValue) >= 50 ? "white" : "black";
    const strength = Math.abs(Number(toneValue) - 50) * 1.6;
    const resolved = `color-mix(in oklch, ${colors.panel} ${100 - strength}%, ${target} ${strength}%)`;
    demo.before = colors.panel;
    demo.after = resolved;
    colors.panel = resolved;
    colors.surface = `color-mix(in oklch, ${colors.surface} ${100 - strength / 2}%, ${target} ${strength / 2}%)`;
    demo.appliedTo = "Chrome and navigation tone";
    demo.expression = `perceptualTone(palette.crust, ${toneValue})`;
    demo.phase = "load/event";
  } else if (functionId === "ensure-contrast") {
    const currentRatio = contrastRatio(colors.selectionText, colors.select);
    const blackRatio = contrastRatio("#000000", colors.select);
    const whiteRatio = contrastRatio("#ffffff", colors.select);
    if (currentRatio < Number(contrastTarget)) colors.selectionText = blackRatio >= whiteRatio ? "#000000" : "#ffffff";
    demo.before = theme.colors.text;
    demo.after = colors.selectionText;
    demo.appliedTo = `Selected-row text · ${contrastRatio(colors.selectionText, colors.select).toFixed(2)}:1`;
    demo.expression = `ensureContrast(app.text, folderView.selection, ${contrastTarget})`;
    demo.phase = "load/event";
  } else if (functionId === "harmonize") {
    const resolved = `color-mix(in oklch, ${colors.accent} ${100 - Number(harmonizeAmount)}%, ${systemColors.accent} ${harmonizeAmount}%)`;
    colors.accent = resolved;
    colors.selectBorder = resolved;
    demo.after = resolved;
    demo.appliedTo = "Accent harmonized to palette target";
    demo.expression = `harmonize(palette.accent, palette.windowsBlue, ${harmonizeAmount}%)`;
    demo.phase = "load";
  }

  return demo;
}

const leftRows = [
  ["folder", "Parent directory", "", ""],
  ["file-text", "00-preview-sample.txt", "2026-07-14 12:41", "1.2 KB · A"],
  ["file-down", "01-release-notes.md", "2026-07-14 12:36", "2.7 KB · A"],
  ["file-text", "02-build-log.log", "2026-07-14 12:32", "84.1 KB · A"],
  ["braces", "03-data.json", "2026-07-14 12:29", "5.3 KB · A"],
  ["file-spreadsheet", "04-report.csv", "2026-07-14 12:24", "40 B · A"],
  ["folder", "Archive Samples", "2026-07-13 17:11", ""],
  ["folder", "Images", "2026-07-13 16:58", ""],
];

const rightRows = [
  ["folder", "Archive Samples", "2026-07-14 12:03", ""],
  ["folder", "Images", "2026-07-14 12:03", ""],
  ["file-text", "00-target-folder.txt", "2026-07-14 12:03", "36 B · A"],
  ["file-spreadsheet", "01-report.csv", "2026-07-14 12:03", "40 B · A"],
  ["image", "02-generated-preview.png", "2026-07-14 12:03", "2.25 KB · A"],
  ["archive", "03-theme-pack.7z", "2026-07-14 11:58", "1.82 MB · A"],
  ["file-code", "04-theme.json5", "2026-07-14 11:51", "7.4 KB · A"],
  ["file-text", "LICENSE.txt", "2026-07-14 11:48", "14.3 KB · A"],
];

const iconMap = {
  archive: Archive,
  braces: Braces,
  check: Check,
  "chevron-down": ChevronDown,
  "chevron-right": ChevronRight,
  copy: Copy,
  "file-code": FileCode,
  "file-down": FileDown,
  "file-spreadsheet": FileSpreadsheet,
  "file-text": FileText,
  folder: Folder,
  "folder-plus": FolderPlus,
  "function-square": FunctionSquare,
  "hard-drive": HardDrive,
  image: Image,
  menu: Menu,
  minus: Minus,
  "move-right": ArrowRight,
  rainbow: Rainbow,
  scale: Scale,
  search: Search,
  sparkles: Sparkles,
  square: Square,
  x: X,
  zap: Zap,
};

function Icon({ name, size = 16, strokeWidth = 1.7 }) {
  const IconComponent = iconMap[name];
  return IconComponent ? <IconComponent width={size} height={size} strokeWidth={strokeWidth} aria-hidden="true" /> : null;
}

function PaletteStrip({ colors }) {
  return <span className="palette-strip" aria-hidden="true">{colors.map((color) => <b key={color} style={{ background: color }} />)}</span>;
}

function ThemeRail({ selectedId, query, onQuery, onSelect }) {
  const visible = themes.filter((theme) => theme.name.toLowerCase().includes(query.toLowerCase()));
  return (
    <aside className="theme-rail">
      <div className="rail-heading">
        <div><span className="eyebrow">Theme catalogue</span><h2>Application preview</h2></div>
        <span className="count-badge">{themes.length}</span>
      </div>
      <label className="theme-search"><Icon name="search" size={14} /><input value={query} onChange={(event) => onQuery(event.target.value)} placeholder="Find a theme" /></label>
      <div className="theme-list" role="listbox" aria-label="Themes">
        {visible.map((theme) => (
          <button key={theme.id} type="button" role="option" aria-selected={selectedId === theme.id} className={`theme-option ${selectedId === theme.id ? "selected" : ""}`} onClick={() => onSelect(theme.id)}>
            <span className="theme-option-copy"><strong>{theme.name}</strong><small>{theme.family}</small></span>
            <PaletteStrip colors={theme.palette} />
            {selectedId === theme.id && <Icon name="check" size={14} />}
          </button>
        ))}
      </div>
      <div className="license-note"><Icon name="scale" size={14} /><span><strong>Palette licenses</strong>Dracula and Catppuccin ship with attribution and MIT license text.</span></div>
    </aside>
  );
}

function Breadcrumb({ active, theme, path, drive = "D:" }) {
  const parts = path.split("\\");
  return (
    <div className={`navigation-view ${active ? "active" : "inactive"}`}>
      <button type="button" className="nav-icon" title="Drive menu"><Icon name="hard-drive" size={17} /></button>
      <button type="button" className="drive-button">{drive}</button>
      {parts.map((part, index) => (
        <span className="crumb" key={`${part}-${index}`}><Icon name="chevron-right" size={14} /><button type="button" style={theme.rainbow ? { borderColor: theme.palette[(index + 1) % theme.palette.length] } : undefined}>{part}</button></span>
      ))}
      <button type="button" className="nav-icon trailing"><Icon name="chevron-down" size={15} /></button>
      <span className="free-space">217 GB</span>
    </div>
  );
}

function FileRow({ row, index, selected, primary, rainbowColor, compact = false }) {
  const [icon, name, modified, size] = row;
  const style = selected && rainbowColor ? { "--row-selection": rainbowColor } : undefined;
  return (
    <div className={`file-row ${selected ? "selected" : ""} ${primary ? "primary" : ""} ${compact ? "compact" : ""}`} style={style}>
      <Icon name={icon} size={compact ? 16 : 18} />
      <span className="file-copy"><strong>{name}</strong>{modified && <small>{modified}{size ? ` · ${size}` : ""}</small>}</span>
    </div>
  );
}

function Pane({ side, theme, active, selectionMode, selectionColors }) {
  const rows = side === "left" ? leftRows : rightRows;
  const selectedIndexes = selectionMode === "single" ? (side === "left" ? [1] : [0]) : (side === "left" ? [1, 3, 4, 5] : [0, 3, 4, 6]);
  const primaryIndex = side === "left" ? 1 : 0;
  return (
    <section className={`file-pane ${active ? "active-pane" : "inactive-pane"}`}>
      <Breadcrumb active={active} theme={theme} path={side === "left" ? "RedSalamander\\.build\\LeftPane" : "RedSalamander\\Docs\\res"} />
      <div className="folder-view">
        {rows.map((row, index) => {
          const selected = selectedIndexes.includes(index);
          const selectionOrder = selectedIndexes.indexOf(index);
          return <FileRow key={row[1]} row={row} index={index} selected={selected} primary={index === primaryIndex} rainbowColor={selected && selectionColors ? selectionColors[(selectionOrder + (side === "right" ? 2 : 0)) % selectionColors.length] : null} compact={side === "left"} />;
        })}
      </div>
      <footer className="pane-status"><span>{selectionMode === "multiple" ? "4 items selected" : "1 item selected"}</span><span>{side === "left" ? "8 items · 95.2 KB" : "8 items · 1.84 MB"}</span></footer>
    </section>
  );
}

function FileMenu({ theme, onClose }) {
  return (
    <div className="file-menu" role="menu">
      <button type="button"><Icon name="copy" size={15} /><span>Copy…</span><kbd>F5</kbd></button>
      <button type="button"><Icon name="move-right" size={15} /><span>Move…</span><kbd>F6</kbd></button>
      <button type="button"><Icon name="folder-plus" size={15} /><span>New folder…</span><kbd>F7</kbd></button>
      <span className="menu-separator" />
      <button type="button"><Icon name="sparkles" size={15} /><span>Compare directories</span></button>
      <button type="button"><Icon name="search" size={15} /><span>Find files…</span><kbd>Alt+F7</kbd></button>
      <span className="menu-separator" />
      <button type="button" onClick={onClose}><Icon name="x" size={15} /><span>Close sample</span></button>
    </div>
  );
}

function ApplicationPreview({ theme, selectionMode, activePane, menuOpen, onMenuToggle, highContrast, seed, phase, demo }) {
  const rainbowOffset = (Number(seed) * 37 + Number(phase)) % 360;
  const rainbowColors = Array.from({ length: 6 }, (_, index) => `hsl(${(rainbowOffset + index * 57) % 360} 72% ${theme.dark ? 28 : 88}%)`);
  const selectionColors = highContrast ? null : (demo.selectionChoices ?? (theme.rainbow ? rainbowColors : null));
  const c = highContrast ? { bg: "#000", surface: "#000", panel: "#000", border: "#fff", text: "#fff", muted: "#fff", accent: "#ffff00", link: "#00ffff", select: "#000080", selectBorder: "#ffff00", selectionText: "#ffffff" } : demo.colors;
  const style = {
    "--app-bg": c.bg,
    "--app-surface": c.surface,
    "--app-panel": c.panel,
    "--app-border": c.border,
    "--app-text": c.text,
    "--app-muted": c.muted,
    "--app-accent": c.accent,
    "--app-link": c.link ?? c.accent,
    "--app-selection": c.select,
    "--app-selection-border": c.selectBorder,
    "--app-selection-text": c.selectionText ?? c.text,
  };
  return (
    <article className={`redsal-window ${theme.dark ? "dark" : "light"} ${theme.mono ? "mono" : ""}`} style={style}>
      <div className="window-titlebar">
        <img src="/redsal-logo.png" alt="" />
        <span>RedSalamander</span>
        <div className="window-actions"><button type="button"><Icon name="minus" size={15} /></button><button type="button"><Icon name="square" size={12} /></button><button type="button"><Icon name="x" size={15} /></button></div>
      </div>
      <nav className="app-menubar" aria-label="Application menu">
        {["Left", "Files", "Edit", "Commands", "Plugins", "View", "Right"].map((item) => <button type="button" className={item === "Files" && menuOpen ? "open" : ""} key={item} onClick={item === "Files" ? onMenuToggle : undefined}>{item}</button>)}
        <button type="button" className="help-menu">Help</button>
      </nav>
      <div className="pane-stage">
        <Pane side="left" theme={theme} active={activePane === "left"} selectionMode={selectionMode} selectionColors={selectionColors} />
        <div className="splitter" aria-hidden="true"><span /><span /><span /></div>
        <Pane side="right" theme={theme} active={activePane === "right"} selectionMode={selectionMode} selectionColors={selectionColors} />
        {menuOpen && <FileMenu theme={theme} onClose={onMenuToggle} />}
      </div>
      <footer className="function-bar">
        {["F1 Help", "F2 Rename", "F3 View", "F4 Edit", "F5 Copy", "F6 Move", "F7 New Folder", "F8 Delete", "F9 Menu", "F10 Quit"].map((item) => <button type="button" key={item}>{item}</button>)}
        <span className="runtime-pill"><Icon name="zap" size={13} /> seed {seed}</span>
      </footer>
    </article>
  );
}

function Segmented({ value, values, onChange, label }) {
  return <div className="segmented" role="group" aria-label={label}>{values.map(([id, text]) => <button type="button" key={id} className={value === id ? "active" : ""} onClick={() => onChange(id)}>{text}</button>)}</div>;
}

function Inspector({ theme, selectionMode, setSelectionMode, activePane, setActivePane, highContrast, setHighContrast, seed, setSeed, phase, setPhase, onRainbowBase, functionId, setFunctionId, demo, systemRole, setSystemRole, toneValue, setToneValue, contrastTarget, setContrastTarget, harmonizeAmount, setHarmonizeAmount }) {
  return (
    <aside className="inspector">
      <section className="theme-summary">
        <span className="eyebrow">Current theme</span>
        <h2>{theme.name}</h2>
        <p>{theme.family} · {theme.base} base</p>
        <PaletteStrip colors={theme.palette} />
        {theme.catppuccin && <div className="cat-mapping" aria-label="Catppuccin semantic mapping"><span><b style={{ background: theme.colors.accent }} />Lavender focus</span><span><b style={{ background: theme.colors.link }} />Blue actions</span><span><b style={{ background: theme.colors.select }} />Overlay2 selection</span><small>Mauve stays available; it is not the global accent.</small></div>}
      </section>
      <section className="function-lab">
        <div className="section-title"><label className="field-label">Dynamic function lab</label><span>{demo.phase}</span></div>
        <div className="function-tabs" role="tablist" aria-label="Theme functions">
          {functionOptions.map(([id, label], index) => <button type="button" key={id} role="tab" aria-selected={functionId === id} className={functionId === id ? "active" : ""} onClick={() => setFunctionId(id)}><b>{index + 1}</b>{label}</button>)}
        </div>
        <div className="function-controls">
          {functionId === "seeded-choice" && <p className="function-hint">Uses stable seed {seed} to choose from the theme's six authored palette references.</p>}
          {functionId === "system-color" && <label className="select-field"><span>Windows role</span><select value={systemRole} onChange={(event) => setSystemRole(event.target.value)}>{Object.keys(systemColors).map((role) => <option key={role} value={role}>{role}</option>)}</select></label>}
          {functionId === "perceptual-tone" && <label className="range-field"><span>OKLCH tone <output>{toneValue}</output></span><input type="range" min="0" max="100" value={toneValue} onInput={(event) => setToneValue(event.currentTarget.value)} onChange={(event) => setToneValue(event.currentTarget.value)} /></label>}
          {functionId === "ensure-contrast" && <><label className="field-label mini-label">Target ratio</label><Segmented label="Contrast target" value={contrastTarget} onChange={setContrastTarget} values={[["3", "3.0"], ["4.5", "4.5"], ["7", "7.0"]]} /></>}
          {functionId === "harmonize" && <label className="range-field"><span>Target influence <output>{harmonizeAmount}%</output></span><input type="range" min="0" max="100" value={harmonizeAmount} onInput={(event) => setHarmonizeAmount(event.currentTarget.value)} onChange={(event) => setHarmonizeAmount(event.currentTarget.value)} /></label>}
        </div>
        <div className="function-result">
          <div className="result-swatches" aria-label="Before and resolved color"><b style={{ background: demo.before }} /><Icon name="chevron-right" size={12} /><b style={{ background: demo.after }} /></div>
          <span><strong>{demo.appliedTo}</strong><code>{demo.expression}</code></span>
        </div>
      </section>
      <section className="view-controls">
        <div><label className="field-label">Selection</label><Segmented label="Selection sample" value={selectionMode} onChange={setSelectionMode} values={[["single", "Single"], ["multiple", "Multiple"]]} /></div>
        <div><label className="field-label">Focused pane</label><Segmented label="Focused pane" value={activePane} onChange={setActivePane} values={[["left", "Left"], ["right", "Right"]]} /></div>
      </section>
      <section className={theme.rainbow ? "rainbow-settings" : "rainbow-settings disabled"}>
        <div className="section-title"><label className="field-label">Rainbow runtime</label>{theme.rainbow && <span>stable</span>}</div>
        <Segmented label="Rainbow base" value={theme.dark ? "dark" : "light"} onChange={onRainbowBase} values={[["light", "Light"], ["dark", "Dark"]]} />
        <label className="range-field"><span>Hue phase <output>{phase}°</output></span><input type="range" min="0" max="360" value={phase} onInput={(event) => setPhase(event.currentTarget.value)} onChange={(event) => setPhase(event.currentTarget.value)} disabled={!theme.rainbow} /></label>
        <label className="number-field"><span>Stable seed</span><input type="number" value={seed} min="0" max="99999" onChange={(event) => setSeed(event.target.value)} disabled={!theme.rainbow} /></label>
      </section>
      <section className="switch-row">
        <span><strong>High contrast</strong><small>Override palette at runtime</small></span>
        <button type="button" className={`switch ${highContrast ? "on" : ""}`} aria-label="High contrast" aria-pressed={highContrast} onClick={() => setHighContrast(!highContrast)}><span /></button>
      </section>
    </aside>
  );
}

export function App() {
  const [themeId, setThemeId] = useState("rainbow-light");
  const [query, setQuery] = useState("");
  const [selectionMode, setSelectionMode] = useState("multiple");
  const [activePane, setActivePane] = useState("left");
  const [menuOpen, setMenuOpen] = useState(false);
  const [highContrast, setHighContrast] = useState(false);
  const [seed, setSeed] = useState(2407);
  const [phase, setPhase] = useState(32);
  const [functionId, setFunctionId] = useState("system-color");
  const [systemRole, setSystemRole] = useState("accent");
  const [toneValue, setToneValue] = useState(68);
  const [contrastTarget, setContrastTarget] = useState("4.5");
  const [harmonizeAmount, setHarmonizeAmount] = useState(42);
  const theme = useMemo(() => themes.find((item) => item.id === themeId) ?? themes[0], [themeId]);
  const demo = useMemo(() => buildFunctionDemo(theme, functionId, { systemRole, toneValue, contrastTarget, harmonizeAmount }), [theme, functionId, systemRole, toneValue, contrastTarget, harmonizeAmount]);

  const selectTheme = (id) => {
    setThemeId(id);
    setHighContrast(false);
    setMenuOpen(false);
  };

  return (
    <main className="theme-workbench">
      <header className="workbench-header">
        <div className="workbench-brand"><img src="/redsal-logo.png" alt="" /><div><h1>RedSalamander Theme Lab</h1><p>Full application preview · no duplicated plugin surfaces</p></div></div>
        <div className="header-actions"><span className="status-dot"><b />Live mockup</span><button type="button" className={menuOpen ? "active" : ""} onClick={() => setMenuOpen(!menuOpen)}><Icon name="menu" size={15} />Menu sample</button></div>
      </header>
      <div className="workbench-body">
        <ThemeRail selectedId={themeId} query={query} onQuery={setQuery} onSelect={selectTheme} />
        <section className="preview-area">
          <div className="preview-heading"><div><span className="eyebrow">Real application anatomy</span><h2>{theme.name} · {selectionMode === "multiple" ? "multiple selection" : "single selection"}</h2></div><div className="preview-key"><span><b className="focus-key" />Current item</span><span><b className="selected-key" />Additional selections</span></div></div>
          <ApplicationPreview theme={theme} selectionMode={selectionMode} activePane={activePane} menuOpen={menuOpen} onMenuToggle={() => setMenuOpen(!menuOpen)} highContrast={highContrast} seed={seed} phase={phase} demo={demo} />
          {selectionMode === "multiple" && <div className="rainbow-explainer"><Icon name={functionId === "seeded-choice" ? "sparkles" : "rainbow"} size={16} /><span><strong>{functionId === "seeded-choice" ? "seededChoice in action" : "Stable multi-selection"}</strong> {functionId === "seeded-choice" ? "Each row selects one authored palette reference from its stable identity." : "Each selected item receives a deterministic tint from its row identity; focus remains visible with the stronger border."}</span></div>}
        </section>
        <Inspector theme={theme} selectionMode={selectionMode} setSelectionMode={setSelectionMode} activePane={activePane} setActivePane={setActivePane} highContrast={highContrast} setHighContrast={setHighContrast} seed={seed} setSeed={setSeed} phase={phase} setPhase={setPhase} onRainbowBase={(base) => setThemeId(base === "dark" ? "rainbow-dark" : "rainbow-light")} functionId={functionId} setFunctionId={setFunctionId} demo={demo} systemRole={systemRole} setSystemRole={setSystemRole} toneValue={toneValue} setToneValue={setToneValue} contrastTarget={contrastTarget} setContrastTarget={setContrastTarget} harmonizeAmount={harmonizeAmount} setHarmonizeAmount={setHarmonizeAmount} />
      </div>
    </main>
  );
}
