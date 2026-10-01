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
// Text areas on a 10 x 7.5 in slide, in points (see the geometry constants in songslides.py).
const TITLE_WIDTH = (10 - 0.5 - 1.3 - 0.4) * 72;
const LYRICS_WIDTH = (10 - 1.0 - 0.4) * 72;
const LYRICS_HEIGHT = (7.5 - 1.4 - 0.8 - 0.1) * 72;
const TOC_TITLE_WIDTH = (10 - 1.0 - 0.4) * 72;
const TOC_COLUMN_WIDTH = ((10 - 1.0 - 0.2) / 2 - 0.2) * 72; // left inset only
const TOC_COLUMN_HEIGHT = (7.5 - 1.6 - 1.2) * 72;

const STORAGE_KEY = "songslides:v1";

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
  privacy: $("privacy"), more: $("more"),
};

// --- parsing and fitting (mirrors songslides.py) --------------------------

function parseSongs(text) {
  const lines = text.replace(/^﻿/, "").replace(/\r\n?/g, "\n").split("\n");
  const songs = [];
  let song = null;
  let stanza = [];
  let ignoredLines = 0;
  const endStanza = () => {
    if (song && stanza.length) song.stanzas.push(stanza);
    stanza = [];
  };
  for (const raw of lines) {
    const line = raw.replace(/\s+$/, "");
    if (line.startsWith("#")) {
      endStanza();
      const title = line.replace(/^#+/, "").trim();
      song = title ? { title, stanzas: [] } : null;
      if (song) songs.push(song);
    } else if (line.trim()) {
      if (song) stanza.push(line);
      else ignoredLines += 1;
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

function songSlide(title, lines, number, count, style) {
  const slide = el("div", "slide");
  slide.setAttribute("role", "img");
  slide.setAttribute("aria-label", `${title}, slide ${number} of ${count}`);
  const titleSize = fitSingleLine(title, TITLE_WIDTH, style.title_size, Math.min(MIN_FITTED_TITLE, style.title_size));
  slide.append(slideTitle(title, titleSize), el("span", "slide-counter", `${number}/${count}`));

  let note = null;
  if (lines.length) {
    const minSize = Math.min(MIN_FITTED_CONTENT, style.content_size);
    const size = fitFontSize(lines, LYRICS_WIDTH, LYRICS_HEIGHT, style.content_size, minSize);
    const list = el("ul", "slide-lyrics");
    list.style.setProperty("--pt", size);
    list.style.setProperty("--gap", Math.round(size * STANZA_GAP_EM));
    for (const line of lines) list.append(el("li", null, line));
    slide.append(list);
    if (textHeight(lines, LYRICS_WIDTH, size, STANZA_GAP_EM) > LYRICS_HEIGHT) {
      note = "Too long to fit. Split this verse with a blank line.";
    } else if (size < style.content_size) {
      note = `Lyrics reduced to ${size} pt to fit`;
    }
  } else {
    note = "No lyrics, so this is a title-only slide";
  }
  const figure = el("figure", "slide-figure");
  figure.append(slide);
  if (note) figure.append(el("p", "slide-note", note));
  return figure;
}

function tocSlides(songs, style) {
  const labels = songs.map((song, i) => `${String(i + 1).padStart(2, " ")}. ${song.title}`);
  const columns = [];
  for (let i = 0; i < labels.length; i += TOC_PER_COLUMN) columns.push(labels.slice(i, i + TOC_PER_COLUMN));
  const size = Math.min(...columns.map((c) =>
    fitFontSize(c, TOC_COLUMN_WIDTH, TOC_COLUMN_HEIGHT, TOC_ENTRY_SIZE, TOC_MIN_ENTRY_SIZE, TOC_GAP_EM)));
  const pages = Math.ceil(labels.length / TOC_PER_SLIDE);
  const figures = [];
  for (let page = 0; page < pages; page++) {
    const heading = "Table of Contents" + (pages > 1 ? ` (${page + 1}/${pages})` : "");
    const slide = el("div", "slide toc");
    slide.setAttribute("role", "img");
    slide.setAttribute("aria-label", heading);
    const list = el("ol", "slide-toc");
    list.style.setProperty("--pt", size);
    list.style.setProperty("--gap", Math.round(size * TOC_GAP_EM));
    for (const label of labels.slice(page * TOC_PER_SLIDE, (page + 1) * TOC_PER_SLIDE)) {
      const item = el("li", null, label);
      item.style.setProperty("--pt", fitSingleLine(label, TOC_COLUMN_WIDTH, size,
        Math.min(TOC_MIN_SINGLE_LINE_SIZE, size), TOC_CHAR_EM));
      list.append(item);
    }
    slide.append(slideTitle(heading, fitSingleLine(heading, TOC_TITLE_WIDTH, style.title_size,
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

// Paint the editor's text behind the transparent textarea, with title lines coloured.
function renderBackdrop() {
  const fragment = document.createDocumentFragment();
  for (const line of els.songs.value.split("\n")) {
    if (line.startsWith("#")) fragment.append(el("span", "title-line", line));
    else fragment.append(line);
    fragment.append("\n");
  }
  fragment.append("\u200b"); // keep a trailing empty line as tall as in the textarea
  els.backdrop.replaceChildren(fragment);
  syncBackdropScroll();
}

function syncBackdropScroll() {
  els.backdrop.scrollTop = els.songs.scrollTop;
}

function renderPreview() {
  const { songs, ignoredLines } = parseSongs(els.songs.value);
  const style = currentStyle();
  els.previewBody.style.setProperty("--font", `"${style.font_family}", Carlito, "Segoe UI", sans-serif`);

  if (!songs.length) {
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
  const tocFigures = els.toc.checked ? tocSlides(songs, style) : [];
  const links = [];
  if (tocFigures.length) {
    nodes.push(group("toc", "Table of contents", plural(tocFigures.length, "slide"), tocFigures));
    links.push(jumpLink("toc", "Contents"));
  }

  let slideCount = tocFigures.length;
  songs.forEach((song, index) => {
    const stanzas = song.stanzas.length ? song.stanzas : [[]];
    slideCount += stanzas.length;
    const figures = stanzas.map((lines, i) => songSlide(song.title, lines, i + 1, stanzas.length, style));
    nodes.push(group(`song-${index + 1}`, song.title, plural(stanzas.length, "slide"), figures));
    links.push(jumpLink(`song-${index + 1}`, song.title));
  });
  els.previewBody.replaceChildren(...nodes);
  els.jump.replaceChildren(...links);
  els.jump.hidden = false;
  showCounts(songs.length, slideCount);
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

async function openSongFile(file) {
  if (!file) return;
  if (!/\.txt$/i.test(file.name) && file.type !== "text/plain") {
    showMessage(`"${file.name}" isn't a .txt file. Save the songs as plain text and open that file.`, "error");
    return;
  }
  els.songs.value = await readSongFile(file);
  renderBackdrop();
  showMessage("");
  renderPreview();
  saveSettings();
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

const worker = new Worker("worker.js");
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

els.songs.addEventListener("input", () => { renderBackdrop(); schedulePreview(); });
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
  if (els.songs.value.trim() && !window.confirm("Replace the songs in the box with the example?")) return;
  els.songs.value = EXAMPLE;
  renderBackdrop();
  renderPreview();
  saveSettings();
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

// On wide screens the controls card is sticky and scrolls inside itself. Size it to the
// space actually visible below the header so the pinned download button is always on screen.
const wideScreen = window.matchMedia("(min-width: 861px)");
function fitControls() {
  if (!wideScreen.matches) {
    els.form.style.maxHeight = "";
    return;
  }
  const top = Math.max(16, els.form.getBoundingClientRect().top);
  els.form.style.maxHeight = `${window.innerHeight - top - 16}px`;
}
window.addEventListener("scroll", fitControls, { passive: true });
window.addEventListener("resize", fitControls);
fitControls();

els.more.open = wideScreen.matches;
restoreSettings();
renderBackdrop();
renderPreview();
