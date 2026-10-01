"""Turn a plain-text song collection into a PowerPoint deck.

This is the single implementation shared by the command line (simple_generator.py)
and the browser app (web/, which runs this file through Pyodide).

Song file format: a line starting with "#" begins a song and holds its title.
Blank lines separate stanzas, and every stanza becomes one slide.
"""

from __future__ import annotations

import io
import math
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Inches, Pt

TITLE_SIZE_RANGE = (16, 72)
CONTENT_SIZE_RANGE = (12, 60)
MIN_FITTED_CONTENT_SIZE = 16  # long stanzas shrink down to this, then may overflow
MIN_FITTED_TITLE_SIZE = 20
COUNTER_SIZE = 24
TOC_ENTRY_SIZE = 20
TOC_MIN_ENTRY_SIZE = 12
TOC_MIN_SINGLE_LINE_SIZE = 14  # a long title shrinks alone, down to this, to stay on one line
TOC_LINE_GAP_EM = 6 / 20
TOC_COLUMNS = 2
TOC_SONGS_PER_COLUMN = 10
TOC_SONGS_PER_SLIDE = TOC_COLUMNS * TOC_SONGS_PER_COLUMN

BLACK = RGBColor(0, 0, 0)
COUNTER_BROWN = RGBColor(139, 69, 19)
LINK_BLUE = RGBColor(0, 0, 139)

# Geometry, tuned to sit below the red line of the church templates.
SIDE_MARGIN = Inches(0.5)
TEXT_INSET = Inches(0.2)
TITLE_TOP = Inches(0.6)
TITLE_HEIGHT = Inches(1.0)
COUNTER_WIDTH = Inches(1.0)
COUNTER_RIGHT = Inches(1.3)  # counter box starts this far from the right edge
CONTENT_TOP = Inches(1.4)
BOTTOM_MARGIN = Inches(0.8)  # clears the footer line of the church templates
TOC_TOP = Inches(1.6)
TOC_BOTTOM_MARGIN = Inches(1.2)
TOC_COLUMN_GAP = Inches(0.2)

# Rough text metrics used to estimate wrapping; Calibri-like fonts average
# about half an em per character.
CHAR_WIDTH_EM = 0.5
BOLD_CHAR_WIDTH_EM = 0.56
TOC_CHAR_WIDTH_EM = 0.42  # measured ~0.41 for Calibri (and metric-identical Carlito) titles
LINE_HEIGHT_EM = 1.2
STANZA_LINE_GAP_EM = 16 / 28  # 16pt after each line at the default 28pt


@dataclass
class Song:
    title: str
    stanzas: list[list[str]] = field(default_factory=list)


@dataclass
class Style:
    font_family: str = "Calibri"
    title_size: int = 32
    content_size: int = 28

    def __post_init__(self):
        self.font_family = (self.font_family or "").strip() or "Calibri"
        self.title_size = _clamp(int(self.title_size), *TITLE_SIZE_RANGE)
        self.content_size = _clamp(int(self.content_size), *CONTENT_SIZE_RANGE)


@dataclass
class Summary:
    songs: int = 0
    song_slides: int = 0
    toc_slides: int = 0
    warnings: list[str] = field(default_factory=list)

    @property
    def total(self) -> int:
        return self.song_slides + self.toc_slides

    def __str__(self) -> str:
        text = f"{self.total} slides from {self.songs} songs"
        if self.toc_slides:
            text += f" (including {self.toc_slides} table of contents slide{'s' if self.toc_slides > 1 else ''})"
        return text

    def to_dict(self) -> dict:
        return {
            "songs": self.songs,
            "song_slides": self.song_slides,
            "toc_slides": self.toc_slides,
            "total": self.total,
            "warnings": list(self.warnings),
        }


# --- parsing ---------------------------------------------------------------

def decode_text(data: bytes) -> str:
    """Decode an uploaded song file: UTF-8 (with or without BOM), else Windows-1252."""
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def parse_songs(text: str) -> list[Song]:
    """Parse "# Title" headed songs; blank lines split stanzas.

    Text before the first title and sections with an empty title are ignored.
    """
    text = text.lstrip("﻿").replace("\r\n", "\n").replace("\r", "\n")
    songs: list[Song] = []
    song: Song | None = None
    stanza: list[str] = []

    def end_stanza():
        if song is not None and stanza:
            song.stanzas.append(stanza.copy())
        stanza.clear()

    for raw in text.split("\n"):
        line = raw.rstrip()
        if line.startswith("#"):
            end_stanza()
            title = line.lstrip("#").strip()
            song = Song(title) if title else None
            if song:
                songs.append(song)
        elif line.strip():
            stanza.append(line)
        else:
            end_stanza()
    end_stanza()
    return songs


# --- text fitting ----------------------------------------------------------

def _wrapped_line_count(text: str, width_pt: float, size: float, char_em: float) -> int:
    chars_per_line = max(1, int(width_pt / (size * char_em)))
    return max(1, math.ceil(len(text) / chars_per_line))


def text_height(lines: list[str], width_pt: float, size: int, gap_em: float = STANZA_LINE_GAP_EM) -> float:
    """Estimated height in points of `lines`, one paragraph each, wrapped to `width_pt`."""
    rows = sum(_wrapped_line_count(line, width_pt, size, CHAR_WIDTH_EM) for line in lines)
    return rows * size * LINE_HEIGHT_EM + len(lines) * size * gap_em


def fit_font_size(lines: list[str], width_pt: float, height_pt: float, size: int, min_size: int,
                  gap_em: float = STANZA_LINE_GAP_EM) -> int:
    """Largest size <= `size` at which the lines fit the box, but never below `min_size`."""
    for candidate in range(size, min_size - 1, -1):
        if text_height(lines, width_pt, candidate, gap_em) <= height_pt:
            return candidate
    return min_size


def fit_single_line(text: str, width_pt: float, size: int, min_size: int,
                    char_em: float = BOLD_CHAR_WIDTH_EM) -> int:
    """Largest size <= `size` at which `text` (bold by default) stays on one line."""
    for candidate in range(size, min_size - 1, -1):
        if _wrapped_line_count(text, width_pt, candidate, char_em) == 1:
            return candidate
    return min_size


# --- building --------------------------------------------------------------

def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _pt(length) -> float:
    return Emu(length).pt


def _remove_all_slides(prs):
    slide_ids = prs.slides._sldIdLst
    for slide_id in list(slide_ids):
        prs.part.drop_rel(slide_id.rId)
        slide_ids.remove(slide_id)


def _pick_layout(prs):
    """Prefer a layout named "Blank", otherwise the one with the fewest placeholders."""
    layouts = list(prs.slide_layouts)
    for layout in layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(layouts, key=lambda layout: len(layout.placeholders))


def _new_slide(prs, layout):
    slide = prs.slides.add_slide(layout)
    # Text goes in our own boxes; drop any empty "Click to add..." placeholders.
    for placeholder in list(slide.placeholders):
        placeholder._element.getparent().remove(placeholder._element)
    return slide


def _text_box(slide, left, top, width, height, inset=TEXT_INSET, wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.margin_left = inset
    frame.margin_right = inset
    frame.vertical_anchor = MSO_ANCHOR.TOP
    frame.word_wrap = wrap
    return frame


def _write(paragraph, text, font_family, size, color, bold=False, align=PP_ALIGN.LEFT, space_after=None):
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_family
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    paragraph.alignment = align
    if space_after is not None:
        paragraph.space_after = Pt(space_after)
    return run


def _add_title(slide, text, style, width):
    frame = _text_box(slide, SIDE_MARGIN, TITLE_TOP, width, TITLE_HEIGHT)
    size = fit_single_line(text, _pt(width - 2 * TEXT_INSET), style.title_size,
                           min(MIN_FITTED_TITLE_SIZE, style.title_size))
    _write(frame.paragraphs[0], text, style.font_family, size, BLACK, bold=True)


def _add_song_slide(prs, layout, song_title, lines, number, count, style, summary):
    slide = _new_slide(prs, layout)
    width = prs.slide_width - SIDE_MARGIN - COUNTER_RIGHT
    _add_title(slide, song_title, style, width)

    counter = _text_box(slide, prs.slide_width - COUNTER_RIGHT, TITLE_TOP, COUNTER_WIDTH, TITLE_HEIGHT,
                        inset=Inches(0.1), wrap=False)
    _write(counter.paragraphs[0], f"{number}/{count}", style.font_family, COUNTER_SIZE,
           COUNTER_BROWN, bold=True, align=PP_ALIGN.RIGHT)

    if lines:
        box_width = prs.slide_width - 2 * SIDE_MARGIN
        box_height = prs.slide_height - CONTENT_TOP - BOTTOM_MARGIN
        frame = _text_box(slide, SIDE_MARGIN, CONTENT_TOP, box_width, box_height)
        text_width = _pt(box_width - 2 * TEXT_INSET)
        text_room = _pt(box_height - Inches(0.1))  # minus the frame's top/bottom insets
        size = fit_font_size(lines, text_width, text_room, style.content_size,
                             min(MIN_FITTED_CONTENT_SIZE, style.content_size))
        if text_height(lines, text_width, size) > text_room:
            summary.warnings.append(f'"{song_title}" slide {number}: stanza is too long to fit; '
                                    "split it with a blank line")
        gap = round(size * STANZA_LINE_GAP_EM)
        for i, line in enumerate(lines):
            paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            _write(paragraph, line, style.font_family, size, BLACK, space_after=gap)
    return slide


def _link_to_slide(run, source_slide, target_slide):
    """Make `run` an internal hyperlink that jumps to `target_slide` when clicked."""
    rId = source_slide.part.relate_to(target_slide.part, RT.SLIDE)
    click = run._r.get_or_add_rPr().add_hlinkClick(rId)
    click.set("action", "ppaction://hlinksldjump")


def _add_toc_slides(prs, layout, entries, style):
    """entries: (title, first slide) per song. Returns the number of TOC slides added."""
    pages = math.ceil(len(entries) / TOC_SONGS_PER_SLIDE)
    column_width = (prs.slide_width - 2 * SIDE_MARGIN - TOC_COLUMN_GAP * (TOC_COLUMNS - 1)) // TOC_COLUMNS
    column_height = prs.slide_height - TOC_TOP - TOC_BOTTOM_MARGIN
    text_width = _pt(column_width - TEXT_INSET)  # TOC columns have no right inset
    labels = [f"{n:2d}. {title}" for n, (title, _) in enumerate(entries, start=1)]
    columns = [labels[i:i + TOC_SONGS_PER_COLUMN] for i in range(0, len(labels), TOC_SONGS_PER_COLUMN)]
    # One entry size for every TOC page: the largest that lets each column fit.
    size = min(fit_font_size(column, text_width, _pt(column_height),
                             TOC_ENTRY_SIZE, TOC_MIN_ENTRY_SIZE, gap_em=TOC_LINE_GAP_EM)
               for column in columns)

    for page in range(pages):
        slide = _new_slide(prs, layout)
        heading = "Table of Contents" + (f" ({page + 1}/{pages})" if pages > 1 else "")
        _add_title(slide, heading, style, prs.slide_width - 2 * SIDE_MARGIN)

        for column in range(TOC_COLUMNS):
            first = page * TOC_SONGS_PER_SLIDE + column * TOC_SONGS_PER_COLUMN
            last = first + TOC_SONGS_PER_COLUMN
            column_entries = list(zip(labels[first:last], entries[first:last]))
            if not column_entries:
                break
            left = SIDE_MARGIN + column * (column_width + TOC_COLUMN_GAP)
            frame = _text_box(slide, left, TOC_TOP, column_width, column_height)
            frame.margin_right = 0
            for i, (label, (_, target)) in enumerate(column_entries):
                paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
                entry_size = fit_single_line(label, text_width, size,
                                             min(TOC_MIN_SINGLE_LINE_SIZE, size), TOC_CHAR_WIDTH_EM)
                run = _write(paragraph, label, style.font_family, entry_size, LINK_BLUE,
                             space_after=round(size * TOC_LINE_GAP_EM))
                _link_to_slide(run, slide, target)
    return pages


def _move_last_slides_to_front(prs, count):
    slide_ids = prs.slides._sldIdLst
    moved = list(slide_ids)[-count:]
    for slide_id in moved:
        slide_ids.remove(slide_id)
    for position, slide_id in enumerate(moved):
        slide_ids.insert(position, slide_id)


def build_presentation(songs, template=None, *, toc=False, style=None):
    """Build the deck. `template` is a .pptx path, file-like object, or None.

    The template's own slides are removed; only its masters, layouts and size are used.
    Returns (Presentation, Summary).
    """
    style = style or Style()
    prs = Presentation(template)
    _remove_all_slides(prs)
    layout = _pick_layout(prs)
    summary = Summary(songs=len(songs))

    entries = []
    for song in songs:
        stanzas = song.stanzas or [[]]
        if not song.stanzas:
            summary.warnings.append(f'"{song.title}" has no lyrics; it gets a title-only slide')
        for number, lines in enumerate(stanzas, start=1):
            slide = _add_song_slide(prs, layout, song.title, lines, number, len(stanzas), style, summary)
            summary.song_slides += 1
            if number == 1:
                entries.append((song.title, slide))

    if toc and entries:
        summary.toc_slides = _add_toc_slides(prs, layout, entries, style)
        _move_last_slides_to_front(prs, summary.toc_slides)
    return prs, summary


def generate_pptx(text, template_bytes=None, *, toc=False, style=None):
    """Song text (+ optional template bytes) -> (.pptx bytes, Summary). Used by the browser app."""
    songs = parse_songs(text)
    if not songs:
        raise ValueError('No songs found. Start each song with a title line such as "# Amazing Grace".')
    template = io.BytesIO(template_bytes) if template_bytes else None
    prs, summary = build_presentation(songs, template, toc=toc, style=style)
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue(), summary
