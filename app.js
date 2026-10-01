"use strict";

// The preview mirrors songslides.py: same parsing rules and the same font-fitting
// estimate, so what you see here is what the .pptx gets. Keep the two in step.

const TITLE_RANGE = [16, 72];
const CONTENT_RANGE = [12, 60];
const MIN_FITTED_CONTENT = 16;
const MIN_FITTED_TITLE = 20;
const TOC_ENTRY_SIZE = 20;
const TOC_MIN_ENTRY_SIZE = 12;
const TOC_MIN_SINGLE_LINE_SIZE = 14;
const TOC_PER_COLUMN = 10;
const TOC_PER_SLIDE = 20;
const CHAR_EM = 0.5;
const BOLD_CHAR_EM = 0.56;
const TOC_CHAR_EM = 0.42;
const LINE_HEIGHT_EM = 1.2;
const STANZA_GAP_EM = 16 / 28;
const TOC_GAP_EM = 6 / 20;
// Free fonts with the same character widths, used when the real font is not installed.
const METRIC_TWINS = { Calibri: "Carlito", Arial: "Arimo", "Times New Roman": "Tinos", Georgia: "Gelasio" };

// Text areas in points for a slide of w x h inches (see the geometry constants in songslides.py).
function geometry(w, h) {
  return {
    w, h,
    titleWidth: (w - 0.5 - 1.3 - 0.4) * 72,
    lyricsWidth: (w - 1.0 - 0.4) * 72,
    lyricsHeight: (h - 1.4 - 0.8 - 0.1) * 72,
    tocTitleWidth: (w - 1.0 - 0.4) * 72,
    tocColumnWidth: ((w - 1.0 - 0.2) / 2 - 0.2) * 72, // left inset only
    tocColumnHeight: (h - 1.6 - 1.2) * 72,
  };
}
const PLAIN_SLIDE = geometry(10, 7.5); // python-pptx's default 4:3 deck

// Wide screens use an editor layout: the song text and the preview scroll separately.
const wideScreen = window.matchMedia("(min-width: 861px)");
const STORAGE_KEY = "songslides:v1";
// The release version, stamped into index.html by tools/publish.sh as app.js?v=…; passing it
// on to the worker and songslides.py means a new release is never mixed with cached old files.
const VERSION = new URL(document.currentScript.src).searchParams.get("v") || "dev";

const EXAMPLE = `# Amazing Grace

Amazing grace! How sweet the sound
That saved a wretch like me!
I once was lost, but now am found;
Was blind, but now I see.

'Twas grace that taught my heart to fear,
And grace my fears relieved;
How precious did that grace appear
The hour I first believed.

Through many dangers, toils and snares,
I have already come;
'Tis grace hath brought me safe thus far,
And grace will lead me home.

# Holy, Holy, Holy

Holy, holy, holy! Lord God Almighty!
Early in the morning our song shall rise to Thee;
Holy, holy, holy, merciful and mighty!
God in three Persons, blessed Trinity!

Holy, holy, holy! All the saints adore Thee,
Casting down their golden crowns around the glassy sea;
Cherubim and seraphim falling down before Thee,
Who wert, and art, and evermore shalt be.
`;

const $ = (id) => document.getElementById(id);
const els = {
  form: $("form"), songs: $("songs"), songFile: $("songFile"), dropZone: $("dropZone"),
  loadExample: $("loadExample"), templateFile: $("templateFile"), templateName: $("templateName"),
  clearTemplate: $("clearTemplate"), toc: $("toc"), fontFamily: $("fontFamily"),
  titleSize: $("titleSize"), contentSize: $("contentSize"), fileName: $("fileName"),
  generate: $("generate"), message: $("message"), engine: $("engine"), engineText: $("engineText"),
  previewBody: $("previewBody"), empty: $("empty"), backdrop: $("backdrop"), jump: $("jump"),
  hymnBoard: $("hymnBoard"), songCount: $("songCount"), slideCount: $("slideCount"),
  templateCard: $("templateCard"), templateSub: $("templateSub"), generateText: $("generateText"),
  privacy: $("privacy"), more: $("more"), search: $("search"), searchCount: $("searchCount"),
  searchPrev: $("searchPrev"), searchNext: $("searchNext"), previewHead: document.querySelector(".preview-head"),
  preview: document.querySelector(".preview"), undo: $("undo"), undoText: $("undoText"), undoButton: $("undoButton"),
};

// --- parsing and fitting (mirrors songslides.py) --------------------------

// Each song also records where its title (titleRange) and verses (ranges) sit in the text,
// as [start, end) character offsets, so a preview slide can select its text in the editor.
function parseSongs(text) {
  const lines = text.replace(/^\ufeff/, "").replace(/\r\n?/g, "\n").split("\n");
  const songs = [];
  let song = null;
  let stanza = [];
  let range = null;
  let ignoredLines = 0;
  let pos = 0;
  const endStanza = () => {
    if (song && stanza.length) {
      song.stanzas.push(stanza);
      song.ranges.push(range);
    }
    stanza = [];
  };
  for (const raw of lines) {
    const start = pos;
    pos += raw.length + 1;
    const line = raw.replace(/\s+$/, "");
    if (line.startsWith("#")) {
      endStanza();
      const title = line.replace(/^#+/, "").trim();
      song = title ? { title, stanzas: [], ranges: [], titleRange: [start, start + line.length] } : null;
      if (song) songs.push(song);
    } else if (line.trim()) {
      if (!song) {
        ignoredLines += 1;
        continue;
      }
      if (!stanza.length) range = [start, start];
      stanza.push(line);
      range[1] = start + line.length;
    } else {
      endStanza();
    }
  }
  endStanza();
  return { songs, ignoredLines };
}

const charCount = (s) => [...s].length;

function wrappedLines(text, width, size, em) {
  const perLine = Math.max(1, Math.floor(width / (size * em)));
  return Math.max(1, Math.ceil(charCount(text) / perLine));
}

function textHeight(lines, width, size, gapEm) {
  const rows = lines.reduce((sum, line) => sum + wrappedLines(line, width, size, CHAR_EM), 0);
  return rows * size * LINE_HEIGHT_EM + lines.length * size * gapEm;
}

function fitFontSize(lines, width, height, size, minSize, gapEm = STANZA_GAP_EM) {
  for (let s = size; s >= minSize; s--) {
    if (textHeight(lines, width, s, gapEm) <= height) return s;
  }
  return minSize;
}

function fitSingleLine(text, width, size, minSize, em = BOLD_CHAR_EM) {
  for (let s = size; s >= minSize; s--) {
    if (wrappedLines(text, width, s, em) === 1) return s;
  }
  return minSize;
}

const clamp = (value, [low, high], fallback) => {
  const n = Number.parseInt(value, 10);
  return Number.isNaN(n) ? fallback : Math.min(high, Math.max(low, n));
};

function currentStyle() {
  return {
    font_family: els.fontFamily.value,
    title_size: clamp(els.titleSize.value, TITLE_RANGE, 32),
    content_size: clamp(els.contentSize.value, CONTENT_RANGE, 28),
  };
}

// --- preview ----------------------------------------------------------------

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function slideTitle(text, size) {
  const title = el("p", "slide-title", text);
  title.style.setProperty("--pt", size);
  return title;
}

// Song slides in text order, with the text range each one came from.
let slideMap = [];

// The template's artwork (background, logo, lines) drawn behind every preview slide.
let templateArt = null; // {svg: SVGElement, urls: [blob URLs], geo}

function newSlide(className) {
  const slide = el("div", className);
  if (templateArt) slide.append(templateArt.svg.cloneNode(true));
  return slide;
}

function songSlide(title, lines, number, count, style, geo, range, from = range[0]) {
  const slide = newSlide("slide");
  slide.setAttribute("role", "button");
  slide.setAttribute("tabindex", "0");
  slide.setAttribute("aria-label", `Edit ${title}, slide ${number} of ${count}`);
  const titleSize = fitSingleLine(title, geo.titleWidth, style.title_size, Math.min(MIN_FITTED_TITLE, style.title_size));
  slide.append(slideTitle(title, titleSize), el("span", "slide-counter", `${number}/${count}`));

  let note = null;
  if (lines.length) {
    const minSize = Math.min(MIN_FITTED_CONTENT, style.content_size);
    const size = fitFontSize(lines, geo.lyricsWidth, geo.lyricsHeight, style.content_size, minSize);
    const list = el("ul", "slide-lyrics");
    list.style.setProperty("--pt", size);
    list.style.setProperty("--gap", Math.round(size * STANZA_GAP_EM));
    for (const line of lines) list.append(el("li", null, line));
    slide.append(list);
    if (textHeight(lines, geo.lyricsWidth, size, STANZA_GAP_EM) > geo.lyricsHeight) {
      note = "Too long to fit. Split this verse with a blank line.";
    } else if (size < style.content_size) {
      note = `Lyrics reduced to ${size} pt to fit`;
    }
  } else {
    note = "No lyrics, so this is a title-only slide";
  }
  const figure = el("figure", "slide-figure");
  figure.dataset.start = range[0];
  figure.dataset.end = range[1];
  figure.append(slide);
  if (note) figure.append(el("p", "slide-note", note));
  // "from": where this slide's text begins; a song's first slide also owns its "# Title" line
  slideMap.push({ from, start: range[0], end: range[1], figure });
  return figure;
}

function tocSlides(songs, style, geo) {
  const labels = songs.map((song, i) => `${String(i + 1).padStart(2, " ")}. ${song.title}`);
  const columns = [];
  for (let i = 0; i < labels.length; i += TOC_PER_COLUMN) columns.push(labels.slice(i, i + TOC_PER_COLUMN));
  const size = Math.min(...columns.map((c) =>
    fitFontSize(c, geo.tocColumnWidth, geo.tocColumnHeight, TOC_ENTRY_SIZE, TOC_MIN_ENTRY_SIZE, TOC_GAP_EM)));
  const pages = Math.ceil(labels.length / TOC_PER_SLIDE);
  const figures = [];
  for (let page = 0; page < pages; page++) {
    const heading = "Table of Contents" + (pages > 1 ? ` (${page + 1}/${pages})` : "");
    const slide = newSlide("slide toc");
    slide.setAttribute("role", "group");
    slide.setAttribute("aria-label", heading);
    const list = el("ol", "slide-toc");
    list.style.setProperty("--pt", size);
    list.style.setProperty("--gap", Math.round(size * TOC_GAP_EM));
    const first = page * TOC_PER_SLIDE;
    labels.slice(first, first + TOC_PER_SLIDE).forEach((label, i) => {
      // like the deck's TOC, each entry jumps to its song
      const link = el("a", null, label);
      link.href = `#song-${first + i + 1}`;
      const item = el("li");
      item.style.setProperty("--pt", fitSingleLine(label, geo.tocColumnWidth, size,
        Math.min(TOC_MIN_SINGLE_LINE_SIZE, size), TOC_CHAR_EM));
      item.append(link);
      list.append(item);
    });
    slide.append(slideTitle(heading, fitSingleLine(heading, geo.tocTitleWidth, style.title_size,
      Math.min(MIN_FITTED_TITLE, style.title_size))), list);
    const figure = el("figure", "slide-figure");
    figure.append(slide);
    figures.push(figure);
  }
  return figures;
}

function group(id, title, detail, figures) {
  const section = el("section", "song");
  section.id = id;
  const head = el("div", "song-head");
  head.append(el("h3", null, title), el("span", null, detail));
  const grid = el("div", "slides");
  grid.append(...figures);
  section.append(head, grid);
  return section;
}

const plural = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;
const PRIVACY = "Runs in this browser tab. Your songs and template are never uploaded.";

function jumpLink(id, label) {
  const link = el("a", null, label);
  link.href = `#${id}`;
  return link;
}

function showCounts(songCount, slideCount) {
  els.songCount.textContent = songCount || "–";
  els.slideCount.textContent = slideCount || "–";
  els.hymnBoard.setAttribute("aria-label", songCount
    ? `${plural(songCount, "song")}, ${plural(slideCount, "slide")}` : "No songs yet");
  els.privacy.textContent = songCount
    ? `${plural(songCount, "song")}, ${plural(slideCount, "slide")}. Nothing is uploaded.` : PRIVACY;
}

// Paint the editor's text behind the transparent textarea: title lines coloured,
// search matches marked. The backdrop's text is exactly the textarea's, so offsets match.
function renderBackdrop() {
  const fragment = document.createDocumentFragment();
  let pos = 0;
  let m = 0;
  for (const line of els.songs.value.split("\n")) {
    const holder = line.startsWith("#") ? el("span", "title-line") : fragment;
    let cursor = 0;
    while (m < search.matches.length && search.matches[m] < pos + line.length) {
      const at = search.matches[m] - pos;
      if (at >= cursor) {
        holder.append(line.slice(cursor, at));
        holder.append(el("mark", m === search.current ? "current" : null, line.slice(at, at + search.term.length)));
        cursor = at + search.term.length;
      }
      m += 1;
    }
    holder.append(line.slice(cursor));
    if (holder !== fragment) fragment.append(holder);
    fragment.append("\n");
    pos += line.length + 1;
  }
  fragment.append("\u200b"); // keep a trailing empty line as tall as in the textarea
  els.backdrop.replaceChildren(fragment);
  syncBackdropScroll();
}

function syncBackdropScroll() {
  els.backdrop.scrollTop = els.songs.scrollTop;
}

// --- editor <-> preview -----------------------------------------------------

// Pixel offset of a character in the editor, measured on the backdrop (same layout):
// drop a zero-width marker at that character and read where it lands.
function editorY(offset) {
  const walker = document.createTreeWalker(els.backdrop, NodeFilter.SHOW_TEXT);
  let seen = 0;
  let node = walker.nextNode();
  let last = node;
  while (node && seen + node.length <= offset) {
    seen += node.length;
    last = node;
    node = walker.nextNode();
  }
  const target = node || last;
  if (!target) return 0;
  const marker = el("span", null, "\u200b");
  const range = document.createRange();
  range.setStart(target, node ? offset - seen : target.length);
  range.insertNode(marker);
  const y = marker.offsetTop;
  const parent = marker.parentNode;
  marker.remove();
  parent.normalize();
  return y;
}

// Scroll the editor so the given character sits a third of the way down.
function scrollEditorTo(offset) {
  els.songs.scrollTop = Math.max(0, editorY(offset) - els.songs.clientHeight / 3);
  syncBackdropScroll();
}

function selectInEditor(start, end) {
  els.songs.focus({ preventScroll: true });
  els.songs.setSelectionRange(start, end);
  scrollEditorTo(start);
}

// The slide the editor is working on, outlined in the preview.
let currentRange = null;

function markCurrentAt(offset, reveal) {
  // the last slide whose text (or its song's title line) starts at or before the offset
  let entry = null;
  for (const item of slideMap) {
    if (item.from > offset) break;
    entry = item;
  }
  entry = entry || slideMap[0];
  for (const figure of els.previewBody.querySelectorAll(".slide-figure.current")) figure.classList.remove("current");
  if (!entry) return;
  currentRange = [entry.start, entry.end];
  entry.figure.classList.add("current");
  if (reveal) revealSlide(entry.figure);
}

// Songs far down the preview render lazily (content-visibility), so a scroll to them first
// lands using estimated heights, and sections drawn a frame later push the target away.
// Re-align every frame until it has stayed put for a few frames (at most ~half a second).
function settleOn(target, block) {
  let previous = null;
  let stable = 0;
  let frames = 0;
  const align = () => {
    target.scrollIntoView({ block });
    const top = Math.round(target.getBoundingClientRect().top);
    stable = top === previous ? stable + 1 : 0;
    previous = top;
    if (stable < 4 && ++frames < 30) requestAnimationFrame(align);
  };
  align();
}

// Bring a slide into the visible part of the preview if it is not already fully in view.
// No smooth gliding: a glide past lazily drawn songs, or one started mid-glide, ends up off.
function revealSlide(figure) {
  // the preview scrolls on its own on wide screens, with the page on narrow ones
  const top = els.previewHead.getBoundingClientRect().bottom;
  const bottom = wideScreen.matches ? els.preview.getBoundingClientRect().bottom : window.innerHeight;
  const box = figure.getBoundingClientRect();
  if (box.top >= top && box.bottom <= bottom) return;
  settleOn(figure, "center");
}

let caretTimer = 0;
function followCaret() {
  clearTimeout(caretTimer);
  caretTimer = setTimeout(() => {
    if (document.activeElement === els.songs) markCurrentAt(els.songs.selectionStart, true);
  }, 120);
}

// --- search -----------------------------------------------------------------

const search = { term: "", matches: [], current: -1 };

function findMatches(keepNear) {
  const term = els.search.value;
  search.term = term;
  search.matches = [];
  if (term) {
    const haystack = els.songs.value.toLowerCase();
    const needle = term.toLowerCase();
    for (let at = haystack.indexOf(needle); at !== -1 && search.matches.length < 5000;
      at = haystack.indexOf(needle, at + needle.length)) {
      search.matches.push(at);
    }
  }
  const near = search.matches.findIndex((at) => at >= keepNear);
  search.current = search.matches.length ? (near === -1 ? 0 : near) : -1;
}

function showSearch(move) {
  const count = search.matches.length;
  els.searchCount.textContent = !search.term ? "" : count ? `${search.current + 1} of ${count}` : "No matches";
  els.searchCount.classList.toggle("none", Boolean(search.term) && !count);
  els.searchPrev.disabled = els.searchNext.disabled = count < 2;
  renderBackdrop();
  if (move && count) {
    const at = search.matches[search.current];
    scrollEditorTo(at);
    els.songs.setSelectionRange(at, at + search.term.length);
    markCurrentAt(at, true);
  }
}

function stepSearch(delta) {
  if (!search.matches.length) return;
  search.current = (search.current + delta + search.matches.length) % search.matches.length;
  showSearch(true);
}

function renderPreview() {
  const { songs, ignoredLines } = parseSongs(els.songs.value);
  const style = currentStyle();
  const geo = templateArt ? templateArt.geo : PLAIN_SLIDE;
  const twin = METRIC_TWINS[style.font_family];
  els.previewBody.style.setProperty("--font", `"${style.font_family}", ${twin ? `"${twin}", ` : ""}sans-serif`);
  els.previewBody.style.setProperty("--slide-w", geo.w);
  els.previewBody.style.setProperty("--slide-ratio", `${geo.w} / ${geo.h}`);
  els.previewBody.style.setProperty("--slide-link", templateArt?.link || "#0000FF");

  if (!songs.length) {
    slideMap = [];
    els.previewBody.replaceChildren(els.empty);
    els.jump.replaceChildren();
    els.jump.hidden = true;
    showCounts(0, 0);
    if (els.songs.value.trim()) {
      els.previewBody.prepend(el("p", "notice", 'No songs found yet. Start each song with a title line such as "# Amazing Grace".'));
    }
    return { songs };
  }

  const nodes = [];
  if (ignoredLines) {
    nodes.push(el("p", "notice",
      `${plural(ignoredLines, "line")} before the first "#" title ${ignoredLines === 1 ? "is" : "are"} left out.`));
  }
  const tocFigures = els.toc.checked ? tocSlides(songs, style, geo) : [];
  const links = [];
  if (tocFigures.length) {
    nodes.push(group("contents", "Table of contents", plural(tocFigures.length, "slide"), tocFigures));
    links.push(jumpLink("contents", "Contents"));
  }

  let slideCount = tocFigures.length;
  slideMap = [];
  songs.forEach((song, index) => {
    const stanzas = song.stanzas.length ? song.stanzas : [[]];
    const ranges = song.stanzas.length ? song.ranges : [song.titleRange];
    slideCount += stanzas.length;
    const figures = stanzas.map((lines, i) =>
      songSlide(song.title, lines, i + 1, stanzas.length, style, geo, ranges[i], i === 0 ? song.titleRange[0] : ranges[i][0]));
    nodes.push(group(`song-${index + 1}`, song.title, plural(stanzas.length, "slide"), figures));
    links.push(jumpLink(`song-${index + 1}`, song.title));
  });
  els.previewBody.replaceChildren(...nodes);
  els.jump.replaceChildren(...links);
  els.jump.hidden = false;
  showCounts(songs.length, slideCount);
  if (currentRange) markCurrentAt(currentRange[0], false);
  return { songs };
}

let renderTimer = 0;
function schedulePreview() {
  clearTimeout(renderTimer);
  renderTimer = setTimeout(() => { renderPreview(); saveSettings(); }, 150);
}

// --- settings that survive a reload (this browser only) ---------------------

function saveSettings() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      text: els.songs.value, toc: els.toc.checked, font: els.fontFamily.value,
      titleSize: els.titleSize.value, contentSize: els.contentSize.value, fileName: els.fileName.value,
    }));
  } catch { /* storage unavailable: nothing to remember */ }
}

function restoreSettings() {
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem(STORAGE_KEY)); } catch { /* ignore */ }
  if (!saved) return;
  if (typeof saved.text === "string") els.songs.value = saved.text;
  if (typeof saved.toc === "boolean") els.toc.checked = saved.toc;
  if ([...els.fontFamily.options].some((o) => o.value === saved.font)) els.fontFamily.value = saved.font;
  if (saved.titleSize) els.titleSize.value = saved.titleSize;
  if (saved.contentSize) els.contentSize.value = saved.contentSize;
  if (saved.fileName) els.fileName.value = saved.fileName;
}

// --- files ------------------------------------------------------------------

async function readSongFile(file) {
  const bytes = await file.arrayBuffer();
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

// Replacing the songs never asks first: it happens at once and can be undone.
let undoText = null;

// After the whole text changes, re-run an open search; otherwise just repaint.
function refreshEditor() {
  if (search.term) {
    findMatches(0);
    showSearch(false);
  } else {
    renderBackdrop();
  }
}

function replaceSongs(text, what) {
  const previous = els.songs.value;
  els.songs.value = text;
  currentRange = null;
  refreshEditor();
  renderPreview();
  saveSettings();
  if (previous.trim() && previous !== text) {
    undoText = previous;
    els.undoText.textContent = `${what} replaced your songs.`;
    els.undo.hidden = false;
  } else {
    hideUndo();
  }
}

function hideUndo() {
  undoText = null;
  els.undo.hidden = true;
}

// Word files are converted by songslides.docx_to_text in the worker; the result lands in
// the editor (with Undo) so titles can be checked in the preview before downloading.
let wordRequest = null;

async function openSongFile(file) {
  if (!file) return;
  if (/\.docx$/i.test(file.name)) {
    wordRequest = { id: `word-${nextId++}`, name: file.name };
    showMessage(engineReady ? `Reading "${file.name}"…` : `"${file.name}" opens when the generator is ready…`);
    worker.postMessage({ type: "word", id: wordRequest.id, document: await file.arrayBuffer() });
    return;
  }
  if (/\.doc$/i.test(file.name)) {
    showMessage(`"${file.name}" is an old Word file. In Word, use File › Save As › Word Document (.docx), then open that.`, "error");
    return;
  }
  if (!/\.txt$/i.test(file.name) && file.type !== "text/plain") {
    showMessage(`"${file.name}" can't be opened. Use a .txt or Word .docx file.`, "error");
    return;
  }
  showMessage("");
  replaceSongs(await readSongFile(file), `"${file.name}"`);
}

function useWordText(text) {
  const { name } = wordRequest;
  wordRequest = null;
  replaceSongs(text, `"${name}"`);
  const count = parseSongs(text).songs.length;
  showMessage(count
    ? `Converted "${name}": ${plural(count, "song")} found. Check the titles in the preview.`
    : `No songs found in "${name}". Add "# " in front of each song title in the box.`, count ? "done" : "error");
}

let templateFile = null;

function setTemplate(file) {
  if (file && !/\.pptx$/i.test(file.name)) {
    showMessage(`"${file.name}" isn't a .pptx file. Choose a PowerPoint template saved as .pptx.`, "error");
    els.templateFile.value = "";
    return;
  }
  templateFile = file || null;
  els.templateName.textContent = templateFile ? templateFile.name : "Choose a template";
  els.templateSub.textContent = templateFile
    ? "Background and logo for every slide" : "Optional .pptx with the church background and logo";
  els.templateCard.classList.toggle("chosen", Boolean(templateFile));
  els.clearTemplate.hidden = !templateFile;
  if (!file) els.templateFile.value = "";
  showMessage("");
  requestTemplateArt(templateFile);
}

// Ask the generator to draw the template's artwork; the preview redraws when it arrives.
let templateRequest = null;

function requestTemplateArt(file) {
  if (templateArt) templateArt.urls.forEach((url) => URL.revokeObjectURL(url));
  templateArt = null;
  templateRequest = file ? `template-${nextId++}` : null;
  renderPreview();
  if (!file) return;
  const id = templateRequest;
  file.arrayBuffer().then((buffer) => worker.postMessage({ type: "template", id, template: buffer }, [buffer]));
}

function useTemplateArt(art) {
  const svg = new DOMParser().parseFromString(art.svg, "image/svg+xml").documentElement;
  if (svg.nodeName !== "svg") throw new Error("not an SVG");
  // Swap the embedded pictures for blob URLs so the many slide copies share one image.
  const urls = [];
  for (const image of svg.querySelectorAll("image")) {
    const match = /^data:([^;,]+);base64,(.*)$/.exec(image.getAttribute("href") || "");
    if (!match) { image.remove(); continue; }
    const bytes = Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0));
    const url = URL.createObjectURL(new Blob([bytes], { type: match[1] }));
    image.setAttribute("href", url);
    urls.push(url);
  }
  svg.setAttribute("class", "slide-art");
  svg.setAttribute("aria-hidden", "true");
  templateArt = {
    svg: document.importNode(svg, true), urls, link: art.link, geo: geometry(art.width_in, art.height_in),
  };
  renderPreview();
}

function outputName() {
  const name = els.fileName.value.trim().replace(/[\\/:*?"<>|]+/g, "_") || "songs_presentation";
  return /\.pptx$/i.test(name) ? name : `${name}.pptx`;
}

function download(bytes, name) {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  const url = URL.createObjectURL(blob);
  const link = el("a");
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

// --- generator worker -------------------------------------------------------

function showMessage(text, kind = "") {
  els.message.textContent = text;
  els.message.className = `message ${kind}`.trim();
}

function setEngine(state, text) {
  els.engine.dataset.state = state;
  els.engineText.textContent = text;
}

const worker = new Worker(`worker.js?v=${VERSION}`);
let engineReady = false;
let pending = null; // {id, name, usedTemplate}
let nextId = 1;

function setBusy(busy) {
  els.generate.disabled = busy || !engineReady;
  els.generateText.textContent = busy ? "Building slides…" : "Download PowerPoint";
}

worker.onmessage = ({ data }) => {
  if (data.type === "status") {
    setEngine("loading", data.text);
  } else if (data.type === "ready") {
    engineReady = true;
    setEngine("ready", "Generator ready");
    setBusy(false);
  } else if (data.type === "result" && pending && data.id === pending.id) {
    download(data.bytes, pending.name);
    const s = data.summary;
    showMessage(`Downloaded ${pending.name}: ${plural(s.total, "slide")} from ${plural(s.songs, "song")}.`, "done");
    pending = null;
    setBusy(false);
  } else if (data.type === "template") {
    if (data.id !== templateRequest) return; // a newer template was chosen meanwhile
    try {
      useTemplateArt(data.art);
    } catch {
      showMessage("The preview can't draw this template's design, but the download still uses it.", "error");
    }
  } else if (data.type === "word") {
    if (wordRequest && data.id === wordRequest.id) useWordText(data.text);
  } else if (data.type === "error" && data.kind === "word") {
    if (!wordRequest || data.id !== wordRequest.id) return;
    wordRequest = null;
    showMessage(data.message, "error");
  } else if (data.type === "error" && data.kind === "template") {
    if (data.id !== templateRequest) return;
    showMessage(/zip|package|content type/i.test(data.message)
      ? "That template couldn't be opened. Save it from PowerPoint as .pptx and choose it again."
      : "The preview can't draw this template's design, but the download still uses it.", "error");
  } else if (data.type === "error") {
    if (data.id === undefined) {
      setEngine("error", "Generator unavailable");
      showMessage("The generator couldn't load. Check your internet connection, then reload the page.", "error");
      return;
    }
    const looksLikeBadTemplate = pending?.usedTemplate && /zip|package|content type/i.test(data.message);
    showMessage(looksLikeBadTemplate
      ? "That template couldn't be opened. Save it from PowerPoint as .pptx and choose it again."
      : data.message, "error");
    pending = null;
    setBusy(false);
  }
};

worker.onerror = () => {
  setEngine("error", "Generator unavailable");
  showMessage("The generator couldn't load. Check your internet connection, then reload the page.", "error");
};

els.form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!engineReady || pending) return;
  const { songs } = renderPreview();
  if (!songs.length) {
    showMessage('No songs found. Start each song with a title line such as "# Amazing Grace".', "error");
    els.songs.focus();
    return;
  }
  const template = templateFile ? await templateFile.arrayBuffer() : null;
  pending = { id: nextId++, name: outputName(), usedTemplate: Boolean(template) };
  setBusy(true);
  showMessage("");
  worker.postMessage({
    type: "generate", id: pending.id, text: els.songs.value, template,
    toc: els.toc.checked, style: currentStyle(),
  }, template ? [template] : []);
});

// --- wiring -----------------------------------------------------------------

els.songs.addEventListener("input", () => {
  hideUndo();
  if (search.term) findMatches(els.songs.selectionStart);
  if (search.term) showSearch(false);
  else renderBackdrop();
  schedulePreview();
  followCaret();
});
els.songs.addEventListener("click", followCaret);
els.songs.addEventListener("keyup", (event) => {
  if (event.key.startsWith("Arrow") || event.key.startsWith("Page") || event.key === "Home" || event.key === "End") followCaret();
});
els.songs.addEventListener("keydown", (event) => {
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "f") {
    event.preventDefault();
    els.search.focus();
    els.search.select();
  }
});

// Search: type to find, Enter / Shift+Enter or the arrows to step, Esc to close.
els.search.addEventListener("input", () => {
  findMatches(els.songs.selectionStart);
  showSearch(true);
});
els.search.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    stepSearch(event.shiftKey ? -1 : 1);
  } else if (event.key === "Escape") {
    const at = search.matches[search.current];
    els.search.value = "";
    findMatches(0);
    showSearch(false);
    if (at !== undefined) selectInEditor(at, at);
  }
});
els.searchNext.addEventListener("click", () => stepSearch(1));
els.searchPrev.addEventListener("click", () => stepSearch(-1));

// Click (or Enter on) a song slide in the preview to select its verse in the editor.
function editSlide(event) {
  const figure = event.target.closest(".slide-figure[data-start]");
  if (!figure || event.target.closest("a")) return;
  if (event.type === "keydown" && event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  const start = Number(figure.dataset.start);
  selectInEditor(start, Number(figure.dataset.end));
  markCurrentAt(start, false);
}
els.previewBody.addEventListener("click", editSlide);
els.previewBody.addEventListener("keydown", editSlide);
els.songs.addEventListener("scroll", syncBackdropScroll);
// The textarea is resizable; keep the backdrop's wrapping width in step.
new ResizeObserver(syncBackdropScroll).observe(els.songs);
for (const input of [els.toc, els.fontFamily, els.titleSize, els.contentSize]) {
  input.addEventListener("input", schedulePreview);
}
els.fileName.addEventListener("change", saveSettings);
els.songFile.addEventListener("change", () => { openSongFile(els.songFile.files[0]); els.songFile.value = ""; });
els.templateFile.addEventListener("change", () => setTemplate(els.templateFile.files[0]));
els.clearTemplate.addEventListener("click", () => setTemplate(null));
els.loadExample.addEventListener("click", () => {
  replaceSongs(EXAMPLE, "The example");
});
els.undoButton.addEventListener("click", () => {
  if (undoText === null) return;
  const text = undoText;
  hideUndo();
  els.songs.value = text;
  refreshEditor();
  renderPreview();
  saveSettings();
  els.songs.focus();
});

// Jump-bar and table-of-contents links: scroll to the song and settle there.
document.addEventListener("click", (event) => {
  const link = event.target.closest('a[href^="#song-"], a[href="#contents"]');
  const target = link && document.getElementById(link.getAttribute("href").slice(1));
  if (!target) return;
  event.preventDefault();
  history.replaceState(null, "", link.getAttribute("href"));
  settleOn(target, "start");
});

for (const type of ["dragenter", "dragover"]) {
  els.dropZone.addEventListener(type, (event) => {
    if (![...event.dataTransfer.types].includes("Files")) return;
    event.preventDefault();
    els.dropZone.classList.add("dragging");
  });
}
for (const type of ["dragleave", "drop"]) {
  els.dropZone.addEventListener(type, () => els.dropZone.classList.remove("dragging"));
}
els.dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;
  event.preventDefault();
  if (/\.pptx$/i.test(file.name)) setTemplate(file);
  else openSongFile(file);
});


restoreSettings();
renderBackdrop();
renderPreview();
