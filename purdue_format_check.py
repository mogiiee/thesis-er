#!/usr/bin/env python3
"""Deterministic Purdue thesis-format checker for a .docx booklet.

Every rule below is taken from `Purdue-standard-thesis-handbook-April-2026.docx`
(section numbers in the rule IDs, e.g. H1.8 = handbook section 1.8) or from the
style, numbering and front-page definitions of
`Purdue-standard-thesis-template-April-2026.docx` (rules prefixed `TPL`). Rules
prefixed `SK` come from the advisor's tracked-change lessons recorded in the
academic-writing skill (caption alignment, table notes, cross-reference order).
`REV` rules cover submission hygiene (no tracked changes or comments).

Every message says what is wrong and what the handbook or template wants
instead, so the report can be acted on without opening this file.

The checker reads the OOXML directly (no Word, no LLM), resolves the effective
paragraph and run properties through the style chain, and reports every
violation with a stable rule ID and a location.

Usage:
    python3 purdue_format_check.py <file.docx>                              # check a document
    python3 purdue_format_check.py <file.docx> --selftest                   # prove each rule fires
    python3 purdue_format_check.py <file.docx> --university purdue          # check against a university config
    python3 purdue_format_check.py <file.docx> --config path/to/uni.json   # check against a config file
"""
from __future__ import annotations

import copy
import json
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W, "wp": WP, "a": A, "m": M}

TNR = "Times New Roman"
TEXT_WIDTH_PT = 468.0  # 8.5in page, 1in margins -> 6.5in = 468pt

MAJOR_HEADINGS = {
    "ACKNOWLEDGMENTS", "ACKNOWLEDGEMENTS", "TABLE OF CONTENTS", "LIST OF TABLES",
    "LIST OF FIGURES", "LIST OF SCHEMES", "LIST OF ABBREVIATIONS", "LIST OF SYMBOLS",
    "NOMENCLATURE", "GLOSSARY", "ABSTRACT", "REFERENCES", "PUBLICATIONS",
    "PUBLICATION", "VITA",
}
FRONT_ZONES = {"TITLE", "COMMITTEE", "DEDICATION"}
LIST_ZONES = {"TABLE OF CONTENTS", "LIST OF TABLES", "LIST OF FIGURES", "LIST OF SCHEMES"}
NO_INDENT_ZONES = {"LIST OF ABBREVIATIONS", "LIST OF SYMBOLS", "NOMENCLATURE", "GLOSSARY"}
SUB_STYLES = {"Heading2": 2, "Heading3": 3, "Heading4": 4, "Heading5": 5,
              "PurdueHeading2": 2, "PurdueHeading3": 3, "PurdueHeading4": 4}
CAP_RE = re.compile(r"^(Table|Figure|Scheme)\s+(\d+)\.(\d+)\.(\s|$)")
NORMAL_HEADINGS = {"ACKNOWLEDGMENTS", "ACKNOWLEDGEMENTS", "TABLE OF CONTENTS"}
BODY_JC = "both"  # the booklet uses justified body text; references must match it
COMMITTEE_HEADERS = {"THE PURDUE UNIVERSITY GRADUATE SCHOOL", "STATEMENT OF COMMITTEE APPROVAL"}  # 14pt in template
MANUAL_LIST_RE = re.compile(r"^(\d{1,2}\.|\(?[a-z]\)|[\u2022\u25cf\-\u2013])\s")
PANEL_RE = re.compile(r"(\(?[a-z]\)\s*)+")
# Heading numbering levels as defined in the template's numbering.xml (abstractNum for Heading 1-5):
# level -> (numFmt, lvlText, ind left, ind hanging). A wrong indent here moves every subheading of that level.
TEMPLATE_HEADING_NUM = {
    1: ("decimal", "%1.", 0, 0),
    2: ("decimal", "%1.%2", 576, 576),
    3: ("decimal", "%1.%2.%3", 720, 720),
    4: ("none", "", 0, 0),
    5: ("none", "", 0, 0),
}
HEADING_LEVEL = {"Heading1": 1, "Heading2": 2, "Heading3": 3, "Heading4": 4, "Heading5": 5}
REVISION_TAGS = {"ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
                 "trPrChange", "tcPrChange", "numberingChange", "tblGridChange"}
TABLE_FITS_PT = 580.0  # 9in text height (648pt) minus a two-line table caption with its 24pt/6pt spacing

# Handbook-derived spacing and indentation constants
BODY_LINE_SPACING = 360  # 1.5 lines (360 twips)
BODY_INDENT_TWIPS = 720  # 0.5 in (720 twips)
FIG_GAP_ABOVE = 480      # 24 pt (480 twips)
TBL_GAP_BELOW = 480      # 24 pt (480 twips)
TABLE_CAP_SPACING = (480, 120)  # 24 pt before, 6 pt after (twips)
FIG_CAP_SPACING = (120, 480)    # 6 pt before, 24 pt after (twips)
CAPTION_LINE_SPACING = 240      # single line spacing (240 twips)
REF_HANGING_INDENT = 720        # 0.5 in (720 twips)
REF_LINE_SPACINGS = (240, 360, 480)  # allowed reference line spacings (twips)


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def local(el) -> str:
    return etree.QName(el).localname


def wval(el, name="val"):
    return el.get(q(name)) if el is not None else None


def onoff(el) -> bool:
    v = wval(el)
    return v not in ("0", "false", "off")


def to_int(v, default=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------- text width
def _font():
    try:
        from PIL import ImageFont
        for path in ("/System/Library/Fonts/Supplemental/Times New Roman.ttf",
                     "/Library/Fonts/Times New Roman.ttf",
                     "C:/Windows/Fonts/times.ttf"):
            try:
                return ImageFont.truetype(path, 12)
            except OSError:
                continue
    except ImportError:
        pass
    return None


_FONT = _font()


def text_width_pt(text: str, size_pt: float = 12.0) -> float:
    """Width of `text` in points at the given size, Times New Roman."""
    if _FONT is not None:
        return _FONT.getlength(text) * size_pt / 12.0
    return len(text) * 0.45 * size_pt  # ponytail: average TNR glyph width when PIL is absent


def line_count(text: str, size_pt: float = 12.0, width_pt: float = TEXT_WIDTH_PT) -> int:
    return 1 if text_width_pt(text, size_pt) <= width_pt else 2


# ---------------------------------------------------------------- accessibility (WCAG 2.1 AA) helpers
GENERIC_ALT_RE = re.compile(r"Description automatically generated|AI-generated content may be incorrect|"
                            r"^A (picture|diagram|graph|chart|screenshot|close-?up|screen shot) (containing|of)\b", re.I)
CONTRAST_MIN = 3.0          # WCAG 1.4.11 non-text contrast for graphical objects
COLOR_SHARE_MIN = 0.01      # colors covering less than 1% of an image (anti-aliasing, thin lines) are ignored
CHART_MAX_COLORS = 32       # a chart/diagram has few significant colors; a photograph or busy screenshot has 50+
# ponytail: a photo crop beside a plain legend can pass as a chart; the remap then shifts a few background tones,
# which is harmless. A real classifier is the upgrade path if that ever matters.
CHROMA_MIN = 40             # max(RGB)-min(RGB) at or above this is a hue that encodes something (bars, segments, zones)
AREA_FILL_SHARE = 0.30      # a color covering more of the image than this is a background region, not an object
# ponytail: greys are skipped on purpose. Grey fills usually sit behind black text, and darkening them to 3:1
# would push that text below 4.5:1. Grey bars at #999999 (2.8:1) therefore pass here; swap them for #666666 by hand.
ABSTRACT_LIMIT = 250       # abstract should aim for this many words or fewer
PAGE_SIZES_TWIPS = ([12240, 15840], [12240, 20160], [15840, 24480])  # letter portrait, legal, and ledger


def _universities_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "universities")


def _as_heading_numbering(raw):
    return {int(k): (v[0], v[1], v[2], v[3]) for k, v in raw.items()}


# config-file key -> (module global, converter). Every value above stays the default, so running
# without --university/--config behaves exactly as before; a university file only overrides keys it sets.
_CONFIG_MAP = {
    "font": ("TNR", str),
    "text_width_pt": ("TEXT_WIDTH_PT", float),
    "table_fits_pt": ("TABLE_FITS_PT", float),
    "page_sizes_twips": ("PAGE_SIZES_TWIPS", lambda v: [list(p) for p in v]),
    "body_alignment": ("BODY_JC", str),
    "abstract_word_limit": ("ABSTRACT_LIMIT", int),
    "major_headings": ("MAJOR_HEADINGS", set),
    "front_zones": ("FRONT_ZONES", set),
    "list_zones": ("LIST_ZONES", set),
    "no_indent_zones": ("NO_INDENT_ZONES", set),
    "normal_headings": ("NORMAL_HEADINGS", set),
    "committee_headers": ("COMMITTEE_HEADERS", set),
    "sub_styles": ("SUB_STYLES", lambda v: {k: int(x) for k, x in v.items()}),
    "heading_level": ("HEADING_LEVEL", lambda v: {k: int(x) for k, x in v.items()}),
    "heading_numbering": ("TEMPLATE_HEADING_NUM", _as_heading_numbering),
    "caption_pattern": ("CAP_RE", re.compile),
    "revision_tags": ("REVISION_TAGS", set),
    "contrast_min": ("CONTRAST_MIN", float),
    "color_share_min": ("COLOR_SHARE_MIN", float),
    "chart_max_colors": ("CHART_MAX_COLORS", int),
    "chroma_min": ("CHROMA_MIN", float),
    "area_fill_share": ("AREA_FILL_SHARE", float),
    "body_line_spacing": ("BODY_LINE_SPACING", int),
    "body_indent_twips": ("BODY_INDENT_TWIPS", int),
    "figure_gap_above_twips": ("FIG_GAP_ABOVE", int),
    "table_gap_below_twips": ("TBL_GAP_BELOW", int),
    "table_caption_spacing": ("TABLE_CAP_SPACING", lambda v: (int(v[0]), int(v[1]))),
    "figure_caption_spacing": ("FIG_CAP_SPACING", lambda v: (int(v[0]), int(v[1]))),
    "caption_line_spacing": ("CAPTION_LINE_SPACING", int),
    "reference_hanging_indent_twips": ("REF_HANGING_INDENT", int),
    "reference_line_spacings": ("REF_LINE_SPACINGS", lambda v: tuple(int(x) for x in v)),
}


def load_university(name_or_path: str) -> dict:
    """Load a university config by name (universities/<name>.json) or by file path."""
    if os.path.isfile(name_or_path):
        path = name_or_path
    else:
        path = os.path.join(_universities_dir(), f"{name_or_path}.json")
    if not os.path.isfile(path):
        known = sorted(f[:-5] for f in os.listdir(_universities_dir()) if f.endswith(".json")) if os.path.isdir(_universities_dir()) else []
        raise FileNotFoundError(f"unknown university {name_or_path!r} ({path}); available: {', '.join(known) or 'none'}")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def apply_config(cfg: dict, ns: dict | None = None) -> None:
    """Rebind module globals from a config dict. Unknown keys raise, so typos fail loudly."""
    target = globals() if ns is None else ns
    for key, value in cfg.items():
        if key in ("school", "handbook", "template"):
            continue  # metadata, not a rule value
        try:
            name, convert = _CONFIG_MAP[key]
        except KeyError:
            raise KeyError(f"unknown config key {key!r}; expected one of: {', '.join(sorted(_CONFIG_MAP))}") from None
        target[name] = convert(value)


def _lum(c):
    def f(u):
        u /= 255.0
        return u / 12.92 if u <= 0.03928 else ((u + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(c[0]) + 0.7152 * f(c[1]) + 0.0722 * f(c[2])


def contrast_ratio(a, b) -> float:
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def image_palette(data: bytes):
    """(is_chart, background, [(rgb, share), ...]) for an image. Palette-like images (charts, diagrams,
    flowcharts) are analysed; photographs and busy screenshots are not, because their colors carry no encoding."""
    import io
    from PIL import Image
    try:
        im = Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        return False, (255, 255, 255), []
    im.thumbnail((256, 256))
    im = im.quantize(64, method=Image.Quantize.MEDIANCUT).convert("RGB")  # merge anti-aliasing shades
    px = list(im.getdata()) if not hasattr(im, "get_flattened_data") else list(im.get_flattened_data())
    w, h = im.size
    counts = Counter(px)
    total = len(px)
    border = Counter(px[:w] + px[-w:] + px[::w] + px[w - 1::w])  # the page shows through at the edges
    bg = border.most_common(1)[0][0]
    palette = [(c, k / total) for c, k in counts.most_common(64) if k / total >= COLOR_SHARE_MIN / 5]
    significant = sum(1 for _, s in palette if s >= COLOR_SHARE_MIN)
    is_chart = significant <= CHART_MAX_COLORS and _lum(bg) > 0.5  # light-background chart or diagram
    return is_chart, bg, palette


def _chromatic(palette):
    """Hue-carrying palette entries that are objects, not the area fill under a curve (see AREA_FILL_SHARE)."""
    return [(c, s) for c, s in palette if max(c) - min(c) >= CHROMA_MIN and s <= AREA_FILL_SHARE]


def low_contrast_colors(data: bytes):
    """Chart colors that fail CONTRAST_MIN against the background: [(rgb, ratio, share)]."""
    is_chart, bg, palette = image_palette(data)
    if not is_chart:
        return []
    return [(c, contrast_ratio(c, bg), s) for c, s in _chromatic(palette)
            if s >= COLOR_SHARE_MIN and contrast_ratio(c, bg) < CONTRAST_MIN]


def _scale_to_luminance(c, target):
    """Scale an RGB color toward black (hue preserved) until its relative luminance is `target`."""
    lo, hi = 0.0, 1.0
    for _ in range(24):
        mid = (lo + hi) / 2
        if _lum(tuple(v * mid for v in c)) >= target:
            hi = mid
        else:
            lo = mid
    return tuple(int(round(v * hi)) for v in c)


def recolor_targets(data: bytes):
    """[(color, replacement)] that lifts every failing chromatic color to CONTRAST_MIN while keeping their
    lightness order, so a sequential colormap (heatmap, confusion matrix) still reads light-to-dark. The failing
    shades are spread between the lightest passing shade and the 3:1 limit rather than all collapsed onto it."""
    is_chart, bg, palette = image_palette(data)
    if not is_chart:
        return []
    goal = CONTRAST_MIN + 0.3  # margin so the checker's quantized re-read still passes
    chrom = _chromatic(palette)
    bad = [c for c, _ in chrom if contrast_ratio(c, bg) < goal]
    if not bad:
        return []
    l_hi = (_lum(bg) + 0.05) / goal - 0.05                       # luminance that just reaches `goal`
    good = [_lum(c) for c, _ in chrom if contrast_ratio(c, bg) >= goal]
    l_lo = min(max(good) + 0.02 if good else l_hi * 0.6, l_hi - 0.01)
    lums = sorted(_lum(c) for c in bad)
    lmin, lmax = lums[0], lums[-1]
    out = []
    for c in bad:
        t = l_hi if lmax == lmin else l_lo + (_lum(c) - lmin) / (lmax - lmin) * (l_hi - l_lo)
        out.append((c, _scale_to_luminance(c, t)))
    return out


def hexcolor(c) -> str:
    return "#%02X%02X%02X" % tuple(c)


def wrapped_lines(text: str, size_pt: float, width_pt: float) -> int:
    """Lines a single-spaced paragraph needs in a column `width_pt` wide (word wrap, no hyphenation)."""
    if width_pt <= 0:
        return 1
    lines, cur = 1, 0.0
    space = text_width_pt(" ", size_pt)
    for word in text.split():
        w = text_width_pt(word, size_pt)
        if cur and cur + space + w > width_pt:
            lines += 1
            cur = w
        else:
            cur = cur + space + w if cur else w
    return lines


# ---------------------------------------------------------------- property parsing
def parse_ppr(ppr) -> dict:
    d: dict = {}
    if ppr is None:
        return d
    for el in ppr:
        t = local(el)
        if t == "pStyle":
            d["pStyle"] = wval(el)
        elif t == "jc":
            d["jc"] = wval(el)
        elif t == "spacing":
            for k in ("before", "after", "line", "lineRule"):
                v = el.get(q(k))
                if v is not None:
                    d[k] = v if k == "lineRule" else to_int(v)
        elif t == "ind":
            for k, dk in (("left", "left"), ("start", "left"), ("firstLine", "firstLine"), ("hanging", "hanging")):
                v = el.get(q(k))
                if v is not None:
                    d[dk] = to_int(v)
        elif t in ("keepNext", "keepLines", "widowControl", "pageBreakBefore", "contextualSpacing"):
            d[t] = onoff(el)
        elif t == "numPr":
            d["numId"] = to_int(wval(el.find("w:numId", NS)))
            d["ilvl"] = to_int(wval(el.find("w:ilvl", NS)), 0)
        elif t == "outlineLvl":
            d["outlineLvl"] = to_int(wval(el))
    return d


def parse_rpr(rpr) -> dict:
    d: dict = {}
    if rpr is None:
        return d
    for el in rpr:
        t = local(el)
        if t == "rStyle":
            d["rStyle"] = wval(el)
        elif t == "rFonts":
            for k in ("ascii", "hAnsi", "asciiTheme", "hAnsiTheme"):
                v = el.get(q(k))
                if v is not None:
                    d[k] = v
        elif t in ("sz", "szCs"):
            d[t] = to_int(wval(el))
        elif t in ("b", "i", "caps", "strike"):
            d[t] = onoff(el)
        elif t == "u":
            d["u"] = wval(el) not in (None, "none")
        elif t == "color":
            d["color"] = wval(el)
        elif t == "vertAlign":
            d["vertAlign"] = wval(el)
    return d


# ---------------------------------------------------------------- document model
@dataclass
class Issue:
    rule: str
    loc: str
    msg: str
    severity: str = "ERROR"  # ERROR = handbook rule; INFO = needs a human eye


@dataclass
class Block:
    idx: int
    el: object
    kind: str                 # p | tbl
    section: int
    zone: str = ""
    text: str = ""
    style: str = ""
    pkind: str = ""           # paragraph classification
    ppr: dict = field(default_factory=dict)
    has_drawing: bool = False
    has_pagebreak: bool = False
    has_sectpr: bool = False


class Styles:
    def __init__(self, styles_xml: bytes, theme_xml: bytes | None):
        self.root = etree.fromstring(styles_xml)
        self.refresh()
        self.theme = {"major": None, "minor": None}
        if theme_xml:
            th = etree.fromstring(theme_xml)
            for key in ("major", "minor"):
                lat = th.find(f".//a:{key}Font/a:latin", NS)
                if lat is not None:
                    self.theme[key] = lat.get("typeface")

    def refresh(self):
        """Re-read the parsed style tables from self.root (call after editing the XML)."""
        root = self.root
        self.ppr: dict[str, dict] = {}
        self.rpr: dict[str, dict] = {}
        self.based: dict[str, str | None] = {}
        self.name: dict[str, str] = {}
        self.type: dict[str, str] = {}
        for s in root.findall("w:style", NS):
            sid = s.get(q("styleId"))
            self.type[sid] = s.get(q("type"))
            self.name[sid] = wval(s.find("w:name", NS)) or sid
            self.based[sid] = wval(s.find("w:basedOn", NS))
            self.ppr[sid] = parse_ppr(s.find("w:pPr", NS))
            self.rpr[sid] = parse_rpr(s.find("w:rPr", NS))
        dd = root.find("w:docDefaults", NS)
        self.def_ppr = parse_ppr(dd.find("w:pPrDefault/w:pPr", NS)) if dd is not None else {}
        self.def_rpr = parse_rpr(dd.find("w:rPrDefault/w:rPr", NS)) if dd is not None else {}

    def chain(self, sid):
        out, seen = [], set()
        while sid and sid in self.ppr and sid not in seen:
            seen.add(sid)
            out.append(sid)
            sid = self.based.get(sid)
        return list(reversed(out))

    def eff_ppr(self, direct: dict) -> dict:
        d = dict(self.def_ppr)
        for sid in self.chain(direct.get("pStyle", "Normal") or "Normal"):
            d.update(self.ppr[sid])
        d.update(direct)
        return d

    def eff_rpr(self, pstyle: str, direct_r: dict) -> dict:
        d = dict(self.def_rpr)
        for sid in self.chain(pstyle or "Normal"):
            d.update(self.rpr[sid])
        for sid in self.chain(direct_r.get("rStyle")):
            d.update(self.rpr[sid])
        d.update(direct_r)
        return d

    def font_of(self, r: dict) -> str | None:
        if r.get("asciiTheme"):
            key = "major" if "major" in r["asciiTheme"] else "minor"
            return self.theme.get(key) or r["asciiTheme"]
        return r.get("ascii")


class Numbering:
    """numbering.xml: numId -> abstractNum -> level, with lvlOverride applied."""

    def __init__(self, xml: bytes | None):
        self.root = etree.fromstring(xml) if xml else None
        self.abstract: dict[int, dict[int, object]] = {}
        self.num: dict[int, tuple[int, dict[int, object]]] = {}
        if self.root is None:
            return
        for an in self.root.findall("w:abstractNum", NS):
            aid = to_int(an.get(q("abstractNumId")))
            self.abstract[aid] = {to_int(l.get(q("ilvl"))): l for l in an.findall("w:lvl", NS)}
        for n in self.root.findall("w:num", NS):
            nid = to_int(n.get(q("numId")))
            overrides = {}
            for ov in n.findall("w:lvlOverride", NS):
                l = ov.find("w:lvl", NS)
                if l is not None:
                    overrides[to_int(ov.get(q("ilvl")))] = l
            self.num[nid] = (to_int(wval(n.find("w:abstractNumId", NS))), overrides)

    def level(self, num_id, ilvl=0):
        """Return the w:lvl element for numId/ilvl, or None."""
        if num_id not in self.num:
            return None
        aid, overrides = self.num[num_id]
        return overrides.get(ilvl) or self.abstract.get(aid, {}).get(ilvl)

    @staticmethod
    def describe(lvl) -> tuple:
        ind = lvl.find("w:pPr/w:ind", NS)
        return (wval(lvl.find("w:numFmt", NS)), wval(lvl.find("w:lvlText", NS)) or "",
                to_int(ind.get(q("left")) if ind is not None else None, 0) or 0,
                to_int(ind.get(q("hanging")) if ind is not None else None, 0) or 0)


class Document:
    def __init__(self, path: str):
        self.path = path
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            self.doc = etree.fromstring(z.read("word/document.xml"))
            self.styles = Styles(z.read("word/styles.xml"),
                                 z.read("word/theme/theme1.xml") if "word/theme/theme1.xml" in names else None)
            self.numbering = Numbering(z.read("word/numbering.xml") if "word/numbering.xml" in names else None)
            self.settings = etree.fromstring(z.read("word/settings.xml")) if "word/settings.xml" in names else None
            self.footnotes = etree.fromstring(z.read("word/footnotes.xml")) if "word/footnotes.xml" in names else None
            self.has_comments = "word/comments.xml" in names
            rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
            self.parts = {}
            self.image_rels: dict[str, str] = {}   # rId -> "word/media/imageN.png"
            for rel in rels:
                tgt = rel.get("Target")
                if tgt and (tgt.startswith("footer") or tgt.startswith("header")):
                    self.parts[rel.get("Id")] = etree.fromstring(z.read("word/" + tgt))
                elif tgt and tgt.startswith("media/"):
                    self.image_rels[rel.get("Id")] = "word/" + tgt
            self.media = {n: z.read(n) for n in names if n.startswith("word/media/")}
        self.media_dirty: set[str] = set()  # media the fixer rewrote and must save
        self._palette_cache: dict[str, list] = {}
        self.body = self.doc.find("w:body", NS)
        self.blocks: list[Block] = []
        self.sects: list = []
        self._walk()

    # -- body flattening (sdt wrappers are transparent)
    def _children(self, el):
        for c in el:
            t = local(c)
            if t == "sdt":
                content = c.find("w:sdtContent", NS)
                if content is not None:
                    yield from self._children(content)
            elif t in ("p", "tbl", "sectPr"):
                yield c

    @staticmethod
    def ptext(p) -> str:
        parts = []
        for el in p.iter():
            t = local(el)
            if t == "t":
                parts.append(el.text or "")
            elif t == "tab":
                parts.append("\t")
            elif t == "delText":
                continue
        return "".join(parts).strip()

    def _walk(self):
        sect = 0
        idx = 0
        pending = []
        for el in self._children(self.body):
            t = local(el)
            if t == "sectPr":
                self.sects.append(el)
                for b in pending:
                    b.section = sect
                break
            b = Block(idx=idx, el=el, kind=t, section=sect)
            idx += 1
            if t == "p":
                b.ppr = parse_ppr(el.find("w:pPr", NS))
                b.style = b.ppr.get("pStyle", "") or ""
                b.text = self.ptext(el)
                b.has_drawing = el.find(".//w:drawing", NS) is not None or el.find(".//w:pict", NS) is not None
                b.has_pagebreak = any(br.get(q("type")) == "page" for br in el.iter(q("br")))
                sp = el.find("w:pPr/w:sectPr", NS)
                if sp is not None:
                    b.has_sectpr = True
                    self.sects.append(sp)
            self.blocks.append(b)
            pending.append(b)
            if b.has_sectpr:
                sect += 1
                pending = []
        self._classify()

    def _classify(self):
        zone = None
        chapter = 0
        for b in self.blocks:
            if b.kind == "tbl":
                b.zone = zone or f"SECT{b.section}"
                continue
            up = re.sub(r"\s+", " ", b.text).upper()
            if b.style == "Heading1":
                chapter += 1
                zone = f"CH{chapter}"
                b.pkind = "heading1"
            elif b.style == "MajorHeading" or b.style == "PurdueMajorHeading":
                zone = up or zone
                b.pkind = "major"
            elif up in MAJOR_HEADINGS and not b.style.startswith("TOC"):
                zone = up
                b.pkind = "major"  # mis-styled major heading; rule H1.12.STYLE catches it
            if zone is None:
                b.zone = {0: "TITLE", 1: "COMMITTEE", 2: "DEDICATION"}.get(b.section, "FRONT")
            else:
                b.zone = zone
            if b.pkind:
                continue
            in_list_zone = b.style.startswith("TOC") or b.style == "TableofFigures"
            if b.style in SUB_STYLES:
                b.pkind = "sub"
            elif b.style == "TableCaption" or (b.text.startswith("Table ") and CAP_RE.match(b.text) and not in_list_zone):
                b.pkind = "tablecap"
            elif b.style == "FigureSchemeCaption" or ((b.text.startswith("Figure ") or b.text.startswith("Scheme ")) and CAP_RE.match(b.text) and not in_list_zone):
                b.pkind = "figcap"
            elif b.style.startswith("TOC") or b.style == "TableofFigures":
                b.pkind = "toc"
            elif b.has_drawing and not b.text:
                b.pkind = "figure"
            elif PANEL_RE.fullmatch(b.text):
                b.pkind = "panel"  # "(a)  (b)" labels under side-by-side figure panels
            elif b.has_pagebreak and not b.text:
                b.pkind = "pagebreak"
            elif not b.text:
                b.pkind = "sectbreak" if b.has_sectpr else "empty"
            elif re.match(r"^Note[:.]", b.text):
                b.pkind = "note"
            elif b.style == "ListParagraph" or b.ppr.get("numId") or MANUAL_LIST_RE.match(b.text):
                b.pkind = "list"
            elif zone == "REFERENCES":
                b.pkind = "ref"
            elif b.style == "BlockQuote":
                b.pkind = "blockquote"
            else:
                b.pkind = "body"

    # -- helpers
    def eff(self, b: Block) -> dict:
        return self.styles.eff_ppr(b.ppr)

    def runs(self, p):
        """Yield (run element, effective rPr, text) for runs with visible text."""
        pstyle = (parse_ppr(p.find("w:pPr", NS)).get("pStyle")) or "Normal"
        for r in p.iter(q("r")):
            if any(local(a) == "del" for a in r.iterancestors()):
                continue
            text = "".join(t.text or "" for t in r.findall("w:t", NS))
            if not text.strip():
                continue
            yield r, self.styles.eff_rpr(pstyle, parse_rpr(r.find("w:rPr", NS))), text

    def table_height_pt(self, tbl) -> float:
        """Estimated printed height of a table: per row, the tallest cell's wrapped single-spaced lines.
        Deterministic given the font metrics; used to decide whether a table can sit on one page."""
        grid = [to_int(g.get(q("w")), 0) or 0 for g in tbl.findall("w:tblGrid/w:gridCol", NS)]
        total = 0.0
        for tr in tbl.findall("w:tr", NS):
            col, row_pt = 0, 0.0
            for tc in tr.findall("w:tc", NS):
                span = to_int(wval(tc.find("w:tcPr/w:gridSpan", NS)), 1) or 1
                width_tw = sum(grid[col:col + span]) if grid else to_int(wval(tc.find("w:tcPr/w:tcW", NS), "w"), 0) or 0
                col += span
                width_pt = max(width_tw / 20.0 - 10.8, 20.0)  # minus the default 0.075in cell margins each side
                cell_pt = 0.0
                for p in tc.iter(q("p")):
                    eff = self.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS)))
                    sizes = [r.get("sz") or 24 for _, r, _ in self.runs(p)] or [24]
                    sz_pt = max(sizes) / 2.0
                    text = self.ptext(p)
                    lines = wrapped_lines(text, sz_pt, width_pt) if text else 1
                    cell_pt += lines * sz_pt * 1.15 + ((eff.get("before") or 0) + (eff.get("after") or 0)) / 20.0
                row_pt = max(row_pt, cell_pt)
            trh = tr.find("w:trPr/w:trHeight", NS)
            if trh is not None:
                row_pt = max(row_pt, (to_int(trh.get(q("val")), 0) or 0) / 20.0)
            total += row_pt
        return total

    def drawing_media(self, drawing) -> str | None:
        """Package path of the picture a <w:drawing> shows, or None for shapes and text boxes."""
        blip = drawing.find(f".//{{{A}}}blip")
        rid = blip.get(f"{{{R}}}embed") if blip is not None else None
        return self.image_rels.get(rid)

    def low_contrast(self, media_name: str):
        if media_name not in self._palette_cache:
            self._palette_cache[media_name] = low_contrast_colors(self.media.get(media_name, b""))
        return self._palette_cache[media_name]

    @staticmethod
    def is_layout_table(tbl) -> bool:
        """A table whose cells hold only pictures and panel labels like '(a)' lays figures out; it carries no data."""
        if tbl.find(".//w:drawing", NS) is None:
            return False
        for tc in tbl.iter(q("tc")):
            txt = Document.ptext(tc)
            if txt and not re.fullmatch(r"\(?[a-z]\)?", txt):
                return False
        return True

    def prev(self, b: Block):
        return self.blocks[b.idx - 1] if b.idx > 0 else None

    def next(self, b: Block):
        return self.blocks[b.idx + 1] if b.idx + 1 < len(self.blocks) else None


# ---------------------------------------------------------------- checker
class Checker:
    def __init__(self, doc: Document):
        self.d = doc
        self.issues: list[Issue] = []

    def add(self, rule, b, msg, severity="ERROR"):
        if isinstance(b, Block):
            loc = f"block {b.idx} [{b.zone}] {b.text[:60]!r}" if b.kind == "p" else f"block {b.idx} [{b.zone}] <table>"
        else:
            loc = str(b)
        self.issues.append(Issue(rule, loc, msg, severity))

    def run(self) -> list[Issue]:
        for name in sorted(n for n in dir(self) if n.startswith("rule_")):
            getattr(self, name)()
        return self.issues

    # ------------------------------------------------ H1.4 margins, page size
    def rule_margins(self):
        d = self.d
        for i, sp in enumerate(d.sects):
            pm = sp.find("w:pgMar", NS)
            ps = sp.find("w:pgSz", NS)
            zones = {b.zone for b in d.blocks if b.section == i}
            front = bool(zones & {"TITLE", "COMMITTEE"})
            want_top = 2160 if front else 1440
            if pm is None:
                self.add("H1.4.MARGIN", f"section {i}", "no page margins defined")
                continue
            got = {k: to_int(pm.get(q(k))) for k in ("top", "left", "right", "bottom")}
            if got["top"] != want_top:
                self.add("H1.4.MARGIN", f"section {i} {sorted(zones)}", f"top margin {got['top']/1440:.2f}in, handbook requires {want_top/1440:.1f}in")
            for k in ("left", "right", "bottom"):
                if got[k] != 1440:
                    self.add("H1.4.MARGIN", f"section {i} {sorted(zones)}", f"{k} margin {got[k]/1440:.2f}in, handbook requires 1.0in")
            if ps is not None:
                w, h = to_int(ps.get(q("w"))), to_int(ps.get(q("h")))
                if sorted((w, h)) not in PAGE_SIZES_TWIPS:
                    self.add("H1.4.PAGESIZE", f"section {i}", f"page size {w}x{h} twips is not Letter, Legal, or Tabloid")
                if ps.get(q("orient")) == "landscape":
                    self.add("H1.7.LANDSCAPE", f"section {i}", "landscape section: page number must be rotated in the left margin (manual check)", "INFO")

    # ------------------------------------------------ H1.6 breaks
    def rule_section_breaks(self):
        for i, sp in enumerate(self.d.sects):
            ty = wval(sp.find("w:type", NS))
            if ty == "continuous":
                self.add("H1.6.CONTINUOUS", f"section {i}", "continuous section break is not allowed")

    def rule_new_page_for_headings(self):
        d = self.d
        for b in d.blocks:
            if b.pkind not in ("heading1", "major"):
                continue
            if b.idx == 0:
                continue
            prev = d.prev(b)
            if prev.kind == "p" and prev.pkind == "empty":
                self.add("H1.12.NOSPACE", b, "empty paragraph directly above a heading (use the style spacing instead)")
            if d.eff(b).get("pageBreakBefore"):
                continue
            starts_new_page = (prev.kind == "p" and (prev.has_pagebreak or prev.has_sectpr)) or \
                any(br.get(q("type")) == "page" for br in b.el.iter(q("br")))
            if not starts_new_page:
                self.add("H1.6.PAGEBREAK", b, "chapter/major heading does not start on a new page (insert a page break before it)")
            elif prev.kind == "p" and prev.pkind == "pagebreak":
                # An empty paragraph that holds only a page break still occupies one line. When the page before
                # it is already full, that line spills onto a fresh page and the break then produces a blank page.
                self.add("H1.6.BLANKPAGE", b, "heading relies on a stand-alone page-break paragraph; this leaves a blank page whenever the previous page is full. Set 'Page break before' on the heading paragraph instead")

    # ------------------------------------------------ H1.7 page numbers
    def rule_page_numbers(self):
        d = self.d
        even_used = d.settings is not None and d.settings.find("w:evenAndOddHeaders", NS) is not None
        for i, sp in enumerate(d.sects):
            refs = {r.get(q("type")): r.get(f"{{{R}}}id") for r in sp if local(r) == "footerReference"}
            zones = {b.zone for b in d.blocks if b.section == i}
            title_sect = "TITLE" in zones
            for typ, rid in refs.items():
                if typ == "even" and not even_used:
                    continue
                part = d.parts.get(rid)
                if part is None:
                    continue
                has_page = any("PAGE" in (t.text or "") and "PAGEREF" not in (t.text or "") for t in part.iter(q("instrText"))) or \
                    any("PAGE" in (f.get(q("instr")) or "") for f in part.iter(q("fldSimple")))
                if title_sect and typ == "first":
                    if has_page:
                        self.add("H1.7.TITLEPAGE", f"section {i} first footer", "title page must not carry a page number")
                    continue
                if not has_page:
                    if typ == "first" and sp.find("w:titlePg", NS) is None:
                        continue  # unused first-page footer
                    self.add("H1.7.PAGENUM", f"section {i} {typ} footer", "footer has no PAGE field")
                    continue
                for p in part.iter(q("p")):
                    if not any("PAGE" in (t.text or "") for t in p.iter(q("instrText"))) and not p.findall(".//w:fldSimple", NS):
                        continue
                    eff = d.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS)))
                    if eff.get("jc") != "center":
                        self.add("H1.7.PAGENUM", f"section {i} {typ} footer", "page number is not centered")
                    for _, r, _ in d.runs(p):
                        if d.styles.font_of(r) != TNR or r.get("sz") != 24:
                            self.add("H1.7.PAGENUM", f"section {i} {typ} footer", f"page number font {d.styles.font_of(r)} {(r.get('sz') or 0)/2}pt, needs Times New Roman 12pt")
                            break
            fmt = sp.find("w:pgNumType", NS)
            if fmt is not None and fmt.get(q("fmt")) not in (None, "decimal"):
                self.add("H1.7.PAGENUM", f"section {i}", f"page number format {fmt.get(q('fmt'))} must be decimal")

    # ------------------------------------------------ H1.8 fonts & sizes
    def _expected_size(self, b: Block):
        if b.pkind in ("heading1", "major"):
            return {28}
        if b.zone == "TITLE":
            return {24, 28}
        if b.zone == "COMMITTEE" and b.text.upper() in COMMITTEE_HEADERS:
            return {28}  # template sets these two header lines at 14pt
        if b.pkind == "note":
            return {20}
        return {24}

    def rule_fonts_and_sizes(self):
        d = self.d
        for b in d.blocks:
            if b.kind == "tbl":
                self._check_table_runs(b)
                continue
            if b.pkind in ("figure", "pagebreak", "empty", "sectbreak"):
                continue
            want = self._expected_size(b)
            seen_sz_issue = False
            for r, eff, text in d.runs(b.el):
                if not re.search(r"[A-Za-z0-9]", text):
                    continue  # symbol-only runs may use a symbol font
                font = d.styles.font_of(eff)
                if font != TNR:
                    self.add("H1.8.FONT", b, f"run {text[:30]!r} uses font {font!r}; all text must be Times New Roman")
                sz = eff.get("sz")
                if sz not in want and not seen_sz_issue and eff.get("vertAlign") not in ("superscript", "subscript"):
                    self.add("H1.8.SIZE", b, f"run {text[:30]!r} is {sz/2 if sz else '?'}pt; this paragraph type must be {'/'.join(str(s/2) for s in sorted(want))}pt")
                    seen_sz_issue = True
                col = eff.get("color")
                if col not in (None, "auto", "000000") and b.pkind in ("body", "list", "ref", "sub", "heading1", "major"):
                    self.add("TPL.COLOR", b, f"run {text[:30]!r} is colored {col}; template text is black")
            if b.pkind == "major" or b.pkind == "heading1":
                if b.text and b.text.upper() != b.text:
                    self.add("H1.12.CAPS", b, "chapter / major heading must be typed in ALL CAPS")

    def _check_table_runs(self, b: Block):
        d = self.d
        for p in b.el.iter(q("p")):
            for r, eff, text in d.runs(p):
                if not re.search(r"[A-Za-z0-9]", text):
                    continue
                font = d.styles.font_of(eff)
                if font != TNR:
                    self.add("H1.8.FONT", b, f"table text {text[:30]!r} uses font {font!r}")
                sz = eff.get("sz")
                if sz is None or not 20 <= sz <= 24:
                    self.add("H1.13.TABLETEXT", b, f"table text {text[:30]!r} is {sz/2 if sz else '?'}pt; must be 10 to 12pt")
                    break

    # ------------------------------------------------ H1.9 / H1.10 / H1.11 body paragraphs
    def rule_body_paragraphs(self):
        d = self.d
        indents = Counter()
        for b in d.blocks:
            if b.kind != "p":
                continue
            eff = d.eff(b)
            line = eff.get("line", 240 if eff.get("lineRule") not in (None, "auto") else None)
            if b.pkind in ("body", "list", "ref", "blockquote") or (b.pkind == "empty" and b.zone.startswith("CH")):
                pass
            if b.pkind == "body" and b.zone not in FRONT_ZONES | LIST_ZONES | NO_INDENT_ZONES:
                if (line, eff.get("lineRule", "auto")) != (BODY_LINE_SPACING, "auto"):
                    self.add("H1.9.SPACING", b, f"line spacing {line}/{eff.get('lineRule')} ; body text must be {BODY_LINE_SPACING/240:.1f} lines ({BODY_LINE_SPACING} auto)")
                if eff.get("jc") not in ("both",):
                    self.add("H1.11.ALIGN", b, f"alignment {eff.get('jc')}; body paragraphs must be justified")
                fl = eff.get("firstLine", 0) or 0
                if eff.get("hanging"):
                    fl = -eff["hanging"]
                if b.zone == "ABSTRACT" and fl == 0:
                    pass  # the abstract is a single un-indented paragraph
                elif fl != BODY_INDENT_TWIPS:
                    self.add("H1.10.INDENT", b, f"first-line indent {fl/1440:.2f}in; indented paragraphs must use {BODY_INDENT_TWIPS/1440:.2f}in")
                indents[fl] += 1
                prev, nxt = d.prev(b), d.next(b)
                before = eff.get("before", 0) or 0
                after = eff.get("after", 0) or 0
                allowed_before = {0}
                if prev is not None and (prev.kind == "tbl" or prev.pkind in ("figure", "panel", "note")):
                    allowed_before = {0, FIG_GAP_ABOVE}
                if before not in allowed_before:
                    self.add("H1.10.PARASPACE", b, f"spacing before {before/20:.0f}pt; indented paragraphs carry no extra space (0pt, or 24pt right after a table/figure)")
                allowed_after = {0}
                if nxt is not None and (nxt.pkind == "figure" or (nxt.kind == "tbl" and nxt.el.find(".//w:drawing", NS) is not None)):
                    allowed_after = {0, FIG_GAP_ABOVE}  # the 24pt gap above a figure that follows a paragraph
                if after not in allowed_after:
                    self.add("H1.10.PARASPACE", b, f"spacing after {after/20:.0f}pt; indented paragraphs carry no extra space")
                if eff.get("left"):
                    self.add("H1.10.INDENT", b, f"left indent {eff['left']/1440:.2f}in on a body paragraph")
            if b.pkind == "empty" and (b.zone.startswith("CH") or b.zone in ("ABSTRACT", "REFERENCES", "ACKNOWLEDGMENTS")):
                self.add("H1.10.EMPTYPARA", b, "empty paragraph used as spacing; remove it and use paragraph spacing")
            if b.pkind == "list" and b.zone.startswith("CH"):
                if (line, eff.get("lineRule", "auto")) != (BODY_LINE_SPACING, "auto"):
                    self.add("H1.9.SPACING", b, f"list item line spacing {line}; must be {BODY_LINE_SPACING/240:.1f} lines")
                if eff.get("jc") != "both":
                    self.add("H1.11.ALIGN", b, f"list item alignment {eff.get('jc')}; must be justified")
            if b.zone in ("TITLE", "COMMITTEE", "DEDICATION", "TABLE OF CONTENTS") and b.pkind not in ("major", "pagebreak", "sectbreak", "figure") and b.text:
                allowed = {360, 480} if b.zone == "COMMITTEE" else {360}  # the template's own committee page uses 480 on department lines
                if line not in allowed or eff.get("lineRule", "auto") != "auto":
                    self.add("H1.9.FRONT", b, f"line spacing {line}; title, committee, dedication, and table of contents pages are 1.5 lines")
            if b.zone == "DEDICATION" and b.pkind == "body":
                if eff.get("jc") != "center":
                    self.add("H1.DEDICATION", b, "dedication must be centered")
                if not all(r.get("i") for _, r, _ in d.runs(b.el)):
                    self.add("H1.DEDICATION", b, "dedication must be italic")
                sp = d.sects[b.section] if b.section < len(d.sects) else None
                if sp is not None and wval(sp.find("w:vAlign", NS)) != "center":
                    self.add("H1.DEDICATION", b, "dedication page must be vertically centered (section vAlign=center)")
        if len([k for k, v in indents.items() if v > 2]) > 1:
            self.add("H1.10.INDENT", "document", f"mixed first-line indents across body paragraphs: {dict(indents)}")

    # ------------------------------------------------ H1.12 headings
    def rule_headings(self):
        d = self.d
        for b in d.blocks:
            if b.kind != "p":
                continue
            eff = d.eff(b)
            up = re.sub(r"\s+", " ", b.text).upper()
            if b.pkind in ("heading1", "major"):
                if b.pkind == "major" and b.style not in ("MajorHeading", "PurdueMajorHeading") and up not in NORMAL_HEADINGS:
                    # ACKNOWLEDGMENTS and TABLE OF CONTENTS are "normal headings" that the template itself formats by hand
                    self.add("H1.12.STYLE", b, f"major heading uses style {b.style or 'Normal'!r}; apply the 'Major Heading' style")
                runs = list(d.runs(b.el))
                if runs and not all(r.get("b") for _, r, _ in runs):
                    self.add("H1.12.HEADFMT", b, "heading must be bold")
                if eff.get("jc") != "center":
                    self.add("H1.12.HEADFMT", b, "heading must be centered")
                if (eff.get("line"), eff.get("lineRule", "auto")) != (240, "auto"):
                    self.add("H1.12.HEADFMT", b, f"heading line spacing {eff.get('line')}; must be single (240)")
                if (eff.get("before", 0) or 0) != 0 or eff.get("after") != 720:
                    self.add("H1.12.HEADFMT", b, f"heading spacing before={eff.get('before', 0)} after={eff.get('after')}; must be 0pt before, 36pt after")
                if eff.get("firstLine") or eff.get("left"):
                    self.add("H1.12.HEADFMT", b, "heading must have no indentation")
                if b.pkind == "heading1" and not eff.get("numId"):
                    self.add("H1.12.NUMBER", b, "chapter heading must be numbered (Heading 1 numbering)")
                if b.pkind == "major" and eff.get("numId"):
                    self.add("H1.12.NUMBER", b, "major heading must not be numbered")
                nxt = d.next(b)
                if nxt is not None and nxt.kind == "p" and nxt.pkind == "empty":
                    self.add("H1.12.NOSPACE", b, "empty paragraph directly below a heading")
            elif b.pkind == "sub":
                lvl = SUB_STYLES[b.style]
                if up == b.text and re.search(r"[A-Za-z]", b.text) and len(b.text) > 3:
                    self.add("H1.12.SENTENCECASE", b, "subheading must be in sentence case, not ALL CAPS")
                if b.text and b.text[0].islower():
                    self.add("H1.12.SENTENCECASE", b, "subheading must start with a capital letter")
                if (eff.get("line"), eff.get("lineRule", "auto")) != (240, "auto"):
                    self.add("H1.12.SUBFMT", b, f"subheading line spacing {eff.get('line')}; must be single")
                if eff.get("before") != 480 or eff.get("after") != 240:
                    self.add("H1.12.SUBFMT", b, f"subheading spacing before={eff.get('before')} after={eff.get('after')}; must be 24pt before, 12pt after")
                for flag in ("keepNext", "keepLines"):
                    if not eff.get(flag, False):
                        self.add("H1.12.SUBFMT", b, f"subheading lacks {flag}")
                if eff.get("widowControl") is False:
                    self.add("H1.12.SUBFMT", b, "subheading has widow/orphan control switched off")
                runs = list(d.runs(b.el))
                bold = all(r.get("b") for _, r, _ in runs) if runs else True
                ital = all(r.get("i") for _, r, _ in runs) if runs else True
                want_bold, want_ital = {2: (True, False), 3: (True, False), 4: (True, True), 5: (False, True)}[lvl]
                if bold != want_bold or ital != want_ital:
                    self.add("H1.12.SUBFMT", b, f"level-{lvl-1} subheading must be {'bold ' if want_bold else ''}{'italic' if want_ital else ''}".strip() + f" (found bold={bold} italic={ital})")
                if lvl == 2 and eff.get("jc") != "center":
                    self.add("H1.12.SUBFMT", b, "first-level subheading must be centered")
                if lvl >= 3 and eff.get("jc") not in (None, "left", "start"):
                    self.add("H1.12.SUBFMT", b, f"subheading alignment {eff.get('jc')}; levels 2 to 4 are left aligned (justified text spreads a wrapped heading across the line)")
                numbered = bool(eff.get("numId"))
                if numbered:
                    lv = d.numbering.level(eff["numId"], eff.get("ilvl", 0))
                    numbered = lv is None or Numbering.describe(lv)[0] != "none"  # the template's level 4/5 exist but print nothing
                if lvl in (2, 3) and not numbered:
                    self.add("H1.12.NUMBER", b, "first- and second-level subheadings must be numbered")
                if lvl in (4, 5) and numbered:
                    self.add("H1.12.NUMBER", b, "third- and fourth-level subheadings must not be numbered")
                nxt, prv = d.next(b), d.prev(b)
                if nxt is not None and nxt.kind == "p" and nxt.pkind == "empty":
                    self.add("H1.12.NOSPACE", b, "empty paragraph directly below a subheading")
                if prv is not None and prv.kind == "p" and prv.pkind == "empty":
                    self.add("H1.12.NOSPACE", b, "empty paragraph directly above a subheading")
                if b.el.find(".//w:footnoteReference", NS) is not None:
                    self.add("H1.18.FOOTNOTE", b, "footnotes are not allowed in headings")
            elif b.pkind == "body" and up in MAJOR_HEADINGS:
                self.add("H1.12.STYLE", b, "major heading typed as a normal paragraph")

    def rule_heading_numbering(self):
        """Template numbering.xml: the heading list levels carry the indent that places each
        heading on the page. A wrong level indent shifts every heading of that level."""
        d = self.d
        seen = set()
        for b in d.blocks:
            if b.kind != "p" or b.style not in HEADING_LEVEL:
                continue
            eff = d.eff(b)
            lvl_no = HEADING_LEVEL[b.style]
            num_id, ilvl = eff.get("numId"), eff.get("ilvl", lvl_no - 1)
            if not num_id:
                continue  # H1.12.NUMBER reports missing numbering
            key = (num_id, ilvl)
            if key in seen:
                continue
            seen.add(key)
            lvl = d.numbering.level(num_id, ilvl)
            if lvl is None:
                self.add("H1.12.NUMIND", f"numbering numId={num_id} ilvl={ilvl} ({b.style})", "heading numbering level is not defined in numbering.xml")
                continue
            fmt, text, left, hanging = Numbering.describe(lvl)
            want_fmt, want_text, want_left, want_hanging = TEMPLATE_HEADING_NUM[lvl_no]
            loc = f"numbering level for {b.style} (numId={num_id} ilvl={ilvl}), e.g. {b.text[:40]!r}"
            if (fmt, text) != (want_fmt, want_text):
                self.add("H1.12.NUMFMT", loc, f"number format {fmt!r} {text!r}; template uses {want_fmt!r} {want_text!r}")
            if (left, hanging) != (want_left, want_hanging):
                self.add("H1.12.NUMIND", loc, f"level indent left={left / 1440:.2f}in hanging={hanging / 1440:.2f}in; template places this heading level at left={want_left / 1440:.2f}in hanging={want_hanging / 1440:.2f}in (every heading of this level is shifted on the page)")

    # ------------------------------------------------ H1.13 captions
    def _cap_number(self, b: Block):
        m = CAP_RE.match(b.text)
        return (m.group(1), int(m.group(2)), int(m.group(3))) if m else None

    def rule_captions(self):
        d = self.d
        chapter = 0
        seq: dict[tuple, int] = defaultdict(int)
        for b in d.blocks:
            if b.pkind == "heading1":
                chapter += 1
            if b.pkind not in ("tablecap", "figcap"):
                continue
            eff = d.eff(b)
            num = self._cap_number(b)
            if not b.text:
                self.add("H1.13.CAPTION", b, "empty caption paragraph")
                continue
            if num is None:
                self.add("H1.13.CAPNUM", b, "caption must start with 'Table N.M.' or 'Figure N.M.' (two numbers)")
            else:
                label, ch, n = num
                if ch != chapter:
                    self.add("H1.13.CAPNUM", b, f"{label} {ch}.{n} sits in chapter {chapter}; chapter number must match")
                seq[(label, ch)] += 1
                if n != seq[(label, ch)]:
                    self.add("H1.13.CAPNUM", b, f"{label} {ch}.{n} out of sequence; expected {label} {ch}.{seq[(label, ch)]}")
                instrs = " ".join((t.text or "") for t in b.el.iter(q("instrText"))) + " ".join(f.get(q("instr")) or "" for f in b.el.iter(q("fldSimple")))
                if f"SEQ {label}" not in instrs or "STYLEREF" not in instrs:
                    self.add("H1.13.INSERTCAP", b, "caption number is typed by hand; use Insert Caption with 'include chapter number'")
            # style and spacing
            if b.pkind == "tablecap":
                if b.style != "TableCaption":
                    self.add("H1.13.CAPSTYLE", b, f"table caption uses style {b.style or 'Normal'!r}; apply 'Table Caption'")
                if eff.get("before") != TABLE_CAP_SPACING[0] or eff.get("after") != TABLE_CAP_SPACING[1]:
                    self.add("H1.13.CAPSPACE", b, f"table caption spacing before={eff.get('before')} after={eff.get('after')}; must be {TABLE_CAP_SPACING[0]//20}pt before, {TABLE_CAP_SPACING[1]//20}pt after")
                nxt = d.next(b)
                if nxt is None or nxt.kind != "tbl":
                    self.add("H1.13.CAPORDER", b, "table caption must sit directly above its table")
                if not eff.get("keepNext"):
                    self.add("H1.13.SAMEPAGE", b, "table caption lacks 'keep with next'; caption and table must stay on one page")
            else:
                if b.style != "FigureSchemeCaption":
                    self.add("H1.13.CAPSTYLE", b, f"figure caption uses style {b.style or 'Normal'!r}; apply 'Figure / Scheme Caption'")
                if eff.get("before") != FIG_CAP_SPACING[0] or eff.get("after") != FIG_CAP_SPACING[1]:
                    self.add("H1.13.CAPSPACE", b, f"figure caption spacing before={eff.get('before')} after={eff.get('after')}; must be {FIG_CAP_SPACING[0]//20}pt before, {FIG_CAP_SPACING[1]//20}pt after")
                prv = d.prev(b)
                ok = prv is not None and (prv.pkind in ("figure", "panel") or (prv.kind == "tbl" and prv.el.find(".//w:drawing", NS) is not None))
                if not ok:
                    self.add("H1.13.CAPORDER", b, "figure caption must sit directly below its picture")
            if (eff.get("line"), eff.get("lineRule", "auto")) != (CAPTION_LINE_SPACING, "auto"):
                self.add("H1.13.CAPSPACE", b, f"caption line spacing {eff.get('line')}; must be single")
            lines = line_count(b.text, 12.0, TEXT_WIDTH_PT - ((eff.get("left") or 0) / 20))
            jc = eff.get("jc")
            if lines == 1 and jc != "center":
                self.add("SK.CAPALIGN", b, f"estimated one-line caption is not centered (found {jc}); confirm in Word", "INFO")
            if lines > 1 and jc != "both":
                self.add("SK.CAPALIGN", b, f"estimated multi-line caption is not justified (found {jc}); confirm in Word", "INFO")
            if any(local(a) in ("tc", "txbxContent") for a in b.el.iterancestors()):
                self.add("H1.13.CAPINTABLE", b, "caption placed inside a table or text box")
            if eff.get("firstLine"):
                self.add("H1.13.CAPSPACE", b, "caption must not carry a first-line indent")
            if not eff.get("keepLines"):
                self.add("H1.13.CAPKEEP", b, "caption lacks 'keep lines together'; a long caption can otherwise split so its second half lands alone on the next page")

    def rule_caption_cross_refs(self):
        d = self.d
        mentions: dict[str, int] = {}
        for b in d.blocks:
            if b.kind != "p" or b.pkind in ("tablecap", "figcap", "toc") or b.zone in LIST_ZONES:
                continue
            for m in re.finditer(r"\b(Table|Figure|Scheme)\s+(\d+\.\d+)(?!\d)", b.text):
                key = f"{m.group(1)} {m.group(2)}"
                mentions.setdefault(key, b.idx)
        for b in d.blocks:
            if b.pkind not in ("tablecap", "figcap"):
                continue
            num = self._cap_number(b)
            if not num:
                continue
            key = f"{num[0]} {num[1]}.{num[2]}"
            if key not in mentions:
                self.add("H1.15.XREF", b, f"{key} is never cross-referenced in the text")
            elif mentions[key] > b.idx:
                self.add("SK.XREFORDER", b, f"{key} is first mentioned after it appears (block {mentions[key]}); refer to it before it appears")

    # ------------------------------------------------ H1.13 figures
    def rule_figures(self):
        d = self.d
        for b in d.blocks:
            for anc in b.el.iter(f"{{{WP}}}anchor"):
                name = anc.find("wp:docPr", NS)
                empty = anc.find(".//w:txbxContent", NS) is not None and not Document.ptext(anc)
                self.add("H1.13.WRAP", b, f"drawing {name.get('name') if name is not None else '?'!r} is floating"
                         + (" and is an empty text box; delete it" if empty else "; Wrap Text must be 'In line with text' so a screen reader meets it where a sighted reader does (WCAG 1.3.2)"))
            for dp in b.el.iter(f"{{{WP}}}docPr"):
                if not (dp.get("descr") or "").strip():
                    self.add("H1.13.ALTTEXT", b, f"drawing {dp.get('name')!r} has no alternative text (WCAG 1.1.1)")
            if b.zone in FRONT_ZONES:
                continue  # the template's own title-page graphics are not thesis figures
            if b.kind == "p" and b.pkind == "panel":
                if not d.eff(b).get("keepNext"):
                    self.add("H1.13.SAMEPAGE", b, "panel label paragraph lacks 'keep with next'; labels and caption must stay with the pictures")
                nxt = d.next(b)
                if nxt is None or nxt.pkind not in ("figcap", "figure"):
                    self.add("H1.13.CAPORDER", b, "panel labels are not followed by the figure caption or the next panel row")
            if b.kind == "p" and b.pkind == "figure":
                eff = d.eff(b)
                if eff.get("jc") != "center":
                    self.add("H1.13.FIGCENTER", b, "figure picture paragraph must be centered")
                if (eff.get("line"), eff.get("lineRule", "auto")) != (240, "auto"):
                    self.add("H1.13.FIGSPACING", b, f"picture paragraph line spacing {eff.get('line')}; must be single")
                if not eff.get("keepNext"):
                    self.add("H1.13.SAMEPAGE", b, "picture paragraph lacks 'keep with next'; picture and caption must stay on one page")
                if b.style in ("Caption", "FigureSchemeCaption", "TableCaption"):
                    self.add("H1.13.CAPSTYLE", b, f"picture paragraph carries caption style {b.style!r}")
                prv = d.prev(b)
                if prv is not None and (prv.kind == "tbl" or prv.pkind in ("body", "list", "note", "sub")):
                    prev_after = d.eff(prv).get("after", 0) if prv.kind == "p" else 0
                    if (eff.get("before", 0) or 0) < FIG_GAP_ABOVE and (prev_after or 0) < FIG_GAP_ABOVE:
                        self.add("H1.13.FIGGAP", b, f"picture following a paragraph or table needs {FIG_GAP_ABOVE//20}pt spacing above it")
                nxt = d.next(b)
                if nxt is None or nxt.pkind != "figcap":
                    if nxt is not None and nxt.pkind in ("figure", "panel"):
                        continue  # stacked panels (or their labels) before one caption
                    self.add("H1.13.CAPORDER", b, "picture has no figure caption directly below it")
                if eff.get("firstLine"):
                    self.add("H1.13.FIGCENTER", b, "picture paragraph carries a first-line indent (breaks centering)")
            if b.kind == "p" and b.has_drawing and b.text and b.pkind != "figure":
                self.add("H1.13.INLINEPARA", b, "picture is embedded inside a text paragraph; place it before or after the paragraph")
        for tx in d.doc.iter(q("txbxContent")):
            txt = Document.ptext(tx)
            if CAP_RE.match(txt or ""):
                self.add("H1.13.CAPINTABLE", "text box", "caption placed inside a text box")
        for b in d.blocks:
            if b.kind != "tbl":
                continue
            for p in b.el.iter(q("p")):
                if CAP_RE.match(Document.ptext(p)):
                    self.add("H1.13.CAPINTABLE", b, f"caption {Document.ptext(p)[:40]!r} placed inside a table cell")

    # ------------------------------------------------ WCAG 2.1 AA (Purdue accessibility requirement)
    def rule_accessibility(self):
        """Headings by style, alt text that says something, no layout tables, 3:1 chart colors,
        no floating objects, no page-sized images of text. Each is checkable from the package itself."""
        d = self.d
        seen_alt: dict[str, str] = {}
        for b in d.blocks:
            # -- headings: bold stand-alone lines are invisible to a screen reader's outline
            if b.kind == "p" and b.pkind == "body" and b.zone.startswith("CH") and b.text and len(b.text) < 90 \
                    and not re.search(r"[.:;,]$", b.text) and len(b.text.split()) <= 12:
                runs = list(d.runs(b.el))
                nxt = d.next(b)
                if runs and all(r.get("b") for _, r, _ in runs) and nxt is not None and nxt.pkind in ("body", "list"):
                    self.add("A11Y.HEADING", b, "bold stand-alone line used as a heading; apply a built-in Heading style (Heading 2-5) so it appears in the screen-reader outline (WCAG 1.3.1)")
            # -- pictures: alt text quality, contrast, scanned pages
            for dr in b.el.iter(q("drawing")):
                dp = dr.find(f".//{{{WP}}}docPr")
                if dp is None:
                    continue
                name, descr = dp.get("name"), (dp.get("descr") or "").strip()
                media = d.drawing_media(dr)
                if descr and GENERIC_ALT_RE.search(descr):
                    self.add("A11Y.ALT", b, f"drawing {name!r} carries Word's automatic alt text {descr[:50]!r}; write what the image conveys (WCAG 1.1.1)")
                elif descr and media:
                    if descr in seen_alt and seen_alt[descr] != media:
                        self.add("A11Y.ALTDUP", b, f"drawing {name!r} repeats the alt text of another picture {descr[:50]!r}; each panel needs its own description")
                    seen_alt.setdefault(descr, media)
                if media and b.zone not in FRONT_ZONES:
                    bad = d.low_contrast(media)
                    if bad:
                        cols = ", ".join(f"{hexcolor(c)} {r:.1f}:1" for c, r, _ in sorted(bad, key=lambda x: -x[2])[:4])
                        self.add("A11Y.CONTRAST", b, f"picture {media.rsplit('/', 1)[-1]} ({name}) uses chart colors below 3:1 against its background: {cols} (WCAG 1.4.11)")
                ext = dr.find(f".//{{{WP}}}extent")
                if ext is not None and int(ext.get("cx", 0)) >= 0.85 * 6.5 * 914400 and int(ext.get("cy", 0)) >= 0.85 * 9 * 914400:
                    nxt = d.next(b)
                    if nxt is None or nxt.pkind != "figcap":
                        self.add("A11Y.SCANNED", b, f"picture {name!r} fills the page and has no caption; a scanned page is not selectable text (WCAG 1.4.5)")
            # -- tables: layout grids read as data tables; merged cells break header association
            if b.kind == "tbl":
                if Document.is_layout_table(b.el):
                    self.add("A11Y.LAYOUTTBL", b, "table holds only pictures and panel labels; a screen reader announces it as a data table with empty headers. Place the panels as inline pictures with the labels in a paragraph below (WCAG 1.3.1)")
                elif b.el.find(".//w:gridSpan", NS) is not None or b.el.find(".//w:vMerge", NS) is not None:
                    self.add("A11Y.TBLMERGE", b, "data table has merged cells; confirm in the PDF checker that every data cell still resolves to its header", "INFO")
        for tx in d.doc.iter(q("txbxContent")):
            if Document.ptext(tx):
                self.add("A11Y.TEXTBOX", "text box", f"text box {Document.ptext(tx)[:40]!r} sits outside the reading order; move its text into the body (WCAG 1.3.2)")

    # ------------------------------------------------ H1.13 tables
    def rule_tables(self):
        d = self.d
        for b in d.blocks:
            if b.kind != "tbl":
                continue
            tblpr = b.el.find("w:tblPr", NS)
            is_fig = b.el.find(".//w:drawing", NS) is not None
            if tblpr is None or wval(tblpr.find("w:jc", NS)) != "center":
                self.add("H1.13.TBLCENTER", b, "table must be horizontally centered (Table Properties > Center)")
            if tblpr is not None and tblpr.find("w:tblpPr", NS) is not None:
                self.add("H1.13.WRAP", b, "table text wrapping must be 'None' (table is floating)")
            justified = 0
            for p in b.el.iter(q("p")):
                eff = d.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS)))
                if (eff.get("line"), eff.get("lineRule", "auto")) != (240, "auto"):
                    self.add("H1.13.TBLSPACING", b, f"table cell text {Document.ptext(p)[:30]!r} is not single spaced")
                    break
                if not is_fig and eff.get("jc") in ("both", "distribute") and Document.ptext(p).strip():
                    justified += 1
            if justified:
                self.add("H1.13.TBLJC", b, f"{justified} cell paragraph(s) are justified (inherited from Normal); in a narrow column Word stretches the word gaps into rivers. Set cell text left-aligned")
            if is_fig:
                nxt = d.next(b)
                if nxt is None or nxt.pkind != "figcap":
                    self.add("H1.13.CAPORDER", b, "multi-panel figure table has no figure caption directly below it")
                prv = d.prev(b)
                if prv is not None and prv.kind == "p" and prv.pkind in ("body", "list", "note") and (d.eff(prv).get("after", 0) or 0) < FIG_GAP_ABOVE:
                    self.add("H1.13.FIGGAP", b, f"figure panels following a paragraph need {FIG_GAP_ABOVE//20}pt spacing above (set {FIG_GAP_ABOVE//20}pt after on the preceding paragraph)")
                for dp in b.el.iter(f"{{{WP}}}docPr"):
                    pass
                continue
            prv = d.prev(b)
            if prv is None or prv.pkind != "tablecap":
                self.add("H1.13.CAPORDER", b, "table has no table caption directly above it")
            nxt = d.next(b)
            if nxt is not None and nxt.kind == "p":
                neff = d.eff(nxt)
                gap_ok = (neff.get("before", 0) or 0) >= TBL_GAP_BELOW or nxt.pkind in ("note", "pagebreak", "sectbreak", "heading1", "major") \
                    or (nxt.pkind == "sub" and (neff.get("before", 0) or 0) >= TBL_GAP_BELOW)
                if not gap_ok:
                    self.add("H1.13.TBLAFTER", b, f"no {TBL_GAP_BELOW//20}pt spacing after the table (next paragraph before={neff.get('before', 0)})")
            rows = b.el.findall("w:tr", NS)
            if nxt is not None and nxt.pkind == "note" and rows and not all(
                    d.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS))).get("keepNext") for p in rows[-1].iter(q("p"))):
                self.add("H1.13.TBLNOTE", b, f"table note {nxt.text[:40]!r} can be pushed to the next page alone; set 'keep with next' on the last row so the note stays under its table")
            for tr in rows:
                for tc in tr.findall("w:tc", NS):
                    if tc.find(".//w:tbl", NS) is not None:
                        self.add("H1.13.NESTED", b, "nested table inside a table")
                        break
            self._check_table_integrity(b, rows)

    def _check_table_integrity(self, b: Block, rows):
        """Handbook 1.13: a table (with caption) begins and ends on one page, and text inside it is 10-12pt.
        Word keeps a table whole only when every row but the last is 'keep with next' and rows cannot split;
        a table that still runs over must repeat its header row so the continuation can be read."""
        d = self.d
        if not rows:
            return
        height = d.table_height_pt(b.el)
        fits = height <= TABLE_FITS_PT

        def kept(tr):
            return all(d.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS))).get("keepNext") for p in tr.iter(q("p")))
        if fits:
            # A table that fits on a page must never be split: chain every row to the next and forbid row breaks.
            if any(tr.find("w:trPr/w:cantSplit", NS) is None for tr in rows):
                self.add("H1.13.TBLKEEP", b, "table rows may split across pages (Table Properties > Row > untick 'Allow row to break across pages')")
            loose = [tr for tr in rows[:-1] if not kept(tr)]
            if loose:
                self.add("H1.13.TBLKEEP", b, f"{len(loose)} row(s) lack 'keep with next'; without it Word splits this {height / 72:.1f}in table across pages even though it fits on one")
        else:
            # Longer than a page: chaining every row makes Word strand the caption and push one row per page.
            # Keep only the header with the first data row and let the header repeat on each continuation.
            if not kept(rows[0]):
                self.add("H1.13.TBLKEEP", b, "header row lacks 'keep with next'; a table longer than a page must at least keep its header with the first data row")
            self.add("H1.13.TBLLONG", b, f"table is about {height / 72:.1f}in tall and cannot sit on one page; it will continue with its header row repeated. Shorten it, or place it on its own page (section break + vertical centering, landscape if wide)", "INFO")
        if rows[0].find("w:trPr/w:tblHeader", NS) is None:
            self.add("H1.13.TBLHEADER", b, "first row is not marked 'Repeat as header row'; a table that runs onto a second page would continue without its header")
        sizes = Counter()
        header_bold = set()
        for ri, tr in enumerate(rows):
            for p in tr.iter(q("p")):
                for _, eff, text in d.runs(p):
                    if not re.search(r"[A-Za-z0-9]", text):
                        continue
                    sizes[eff.get("sz")] += 1
                    if ri == 0:
                        header_bold.add(bool(eff.get("b")))
        if len(sizes) > 1:
            self.add("H1.13.TBLFONTMIX", b, f"table mixes font sizes {sorted((s or 0) / 2 for s in sizes)}pt; use one size (10 to 12pt) throughout the table")
        if len(header_bold) > 1:
            self.add("H1.13.TBLHEADFMT", b, "header row mixes bold and regular cells; format every header cell the same way")

    def rule_table_notes(self):
        d = self.d
        for b in d.blocks:
            if b.pkind != "note":
                continue
            eff = d.eff(b)
            runs = list(d.runs(b.el))
            if not all(r.get("i") for _, r, _ in runs):
                self.add("H1.13.NOTE", b, "table/figure note must be italic")
            if (eff.get("line"), eff.get("lineRule", "auto")) != (240, "auto"):
                self.add("H1.13.NOTE", b, "note must be single spaced")
            if eff.get("jc") not in ("left", "start", "both", None):
                self.add("H1.13.NOTE", b, f"note alignment {eff.get('jc')}; must be left or justified (handbook text says left, its examples say justified)")
            # Handbook text: 6pt before / 24pt after. Its own examples: 0pt before, 24pt after (table), 12pt after (figure).
            prv = d.prev(b)
            after_ok = {480} | ({240} if prv is not None and prv.pkind == "figcap" else set())
            if (eff.get("before", 0) or 0) not in (0, 120) or eff.get("after") not in after_ok:
                self.add("H1.13.NOTE", b, f"note spacing before={eff.get('before', 0)} after={eff.get('after')}; must be 0 or 6pt before, {'12 or ' if 240 in after_ok else ''}24pt after")
            if eff.get("firstLine"):
                self.add("H1.13.NOTE", b, "note must not be indented")

    # ------------------------------------------------ H1.16 equations, H1.17 block quotes
    def rule_equations(self):
        d = self.d
        for om in d.doc.iter(f"{{{M}}}oMath"):
            p = om
            while p is not None and local(p) != "p":
                p = p.getparent()
            if p is None:
                continue
            txt = Document.ptext(p)
            if not re.search(r"\(\d+\.\d+\)\s*$", txt):
                self.add("H1.16.EQNUM", f"equation {txt[:40]!r}", "equation must be numbered (chapter.number) at the right margin")
            eff = d.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS)))
            if eff.get("jc") != "right":
                self.add("H1.16.EQALIGN", f"equation {txt[:40]!r}", "equation paragraph must be right aligned with tabs centering the equation")
            if (eff.get("line"), eff.get("lineRule", "auto")) != (480, "auto"):
                self.add("H1.16.EQSPACING", f"equation {txt[:40]!r}", "equations use double line spacing")
            for r in om.iter(f"{{{M}}}r"):
                rpr = parse_rpr(r.find("w:rPr", NS))
                if rpr.get("sz") not in (None, 24):
                    self.add("H1.16.EQFONT", f"equation {txt[:40]!r}", "equation font size must be 12pt")

    def rule_block_quotes(self):
        d = self.d
        for b in d.blocks:
            if b.pkind != "blockquote":
                continue
            eff = d.eff(b)
            if eff.get("left") != 720:
                self.add("H1.17.BLOCKQUOTE", b, "block quote must have a 0.5in left indent")
            if eff.get("jc") != "both":
                self.add("H1.17.BLOCKQUOTE", b, "block quote must be justified")
            if eff.get("before") != 120 or eff.get("after") != 240:
                self.add("H1.17.BLOCKQUOTE", b, "block quote spacing must be 6pt before, 12pt after")
            if b.text.startswith(('"', "\u201c")):
                self.add("H1.17.BLOCKQUOTE", b, "block quotes carry no quotation marks")

    # ------------------------------------------------ H1.18 footnotes
    def rule_footnotes(self):
        d = self.d
        if d.footnotes is None:
            return
        for fn in d.footnotes.findall("w:footnote", NS):
            if to_int(fn.get(q("id")), 0) <= 0:
                continue
            for p in fn.findall("w:p", NS):
                eff = d.styles.eff_ppr(parse_ppr(p.find("w:pPr", NS)))
                txt = Document.ptext(p)
                if (eff.get("line"), eff.get("lineRule", "auto")) != (240, "auto") or (eff.get("before", 0) or 0) != 0 or (eff.get("after", 0) or 0) != 0:
                    self.add("H1.18.FOOTNOTE", f"footnote {txt[:30]!r}", "footnote must be single spaced with 0pt before and after")
                if eff.get("jc") not in ("left", "start", "both", None):
                    self.add("H1.18.FOOTNOTE", f"footnote {txt[:30]!r}", "footnote must be left or justified")
                if not eff.get("keepLines", False) or eff.get("widowControl") is False:
                    self.add("H1.18.FOOTNOTE", f"footnote {txt[:30]!r}", "footnote needs widow/orphan control and keep lines together")
                for _, r, text in d.runs(p):
                    if re.search(r"[A-Za-z]", text) and (d.styles.font_of(r) != TNR or r.get("sz") != 20):
                        self.add("H1.18.FOOTNOTE", f"footnote {txt[:30]!r}", "footnote text must be Times New Roman 10pt")
                        break
        if d.settings is not None:
            fp = d.settings.find("w:footnotePr", NS)
            if fp is not None:
                if wval(fp.find("w:numFmt", NS)) not in (None, "decimal"):
                    self.add("H1.18.FOOTNOTE", "settings", "footnote numbering must use Arabic numerals")
                if wval(fp.find("w:numRestart", NS)) not in (None, "continuous"):
                    self.add("H1.18.FOOTNOTE", "settings", "footnote numbering must be continuous through the document")
        for sp in d.sects:
            fp = sp.find("w:footnotePr", NS)
            if fp is not None and wval(fp.find("w:numRestart", NS)) not in (None, "continuous"):
                self.add("H1.18.FOOTNOTE", "section", "footnote numbering restarts in a section")

    # ------------------------------------------------ H1.21 references
    def rule_references(self):
        d = self.d
        expected = 0
        for b in d.blocks:
            if b.zone != "REFERENCES" or b.kind != "p" or b.pkind in ("major", "pagebreak", "sectbreak"):
                continue
            if b.pkind == "empty":
                continue
            eff = d.eff(b)
            m = re.match(r"^\[(\d+)\]", b.text)
            if not m:
                self.add("H1.21.REFNUM", b, "reference entry does not start with a [n] number")
            else:
                expected += 1
                if int(m.group(1)) != expected:
                    self.add("H1.21.REFNUM", b, f"reference numbered [{m.group(1)}], expected [{expected}] (numerical order)")
            if eff.get("hanging") != REF_HANGING_INDENT:
                self.add("H1.21.REFHANG", b, f"reference needs a {REF_HANGING_INDENT/1440:.2f}in hanging indent (found {eff.get('hanging')})")
            line = eff.get("line")
            if line not in REF_LINE_SPACINGS:
                self.add("H1.21.REFSPACING", b, f"reference line spacing {line}; must be 1.0, 1.5, or 2.0")
            before, after = eff.get("before", 0) or 0, eff.get("after", 0) or 0
            # single: 0 before, at most 12pt after; 1.5: 0 before, exactly 6pt after
            spacing_ok = before == 0 and ((line == 240 and after <= 240) or (line == 360 and after == 120) or (line == 480 and after == 0))
            if not spacing_ok:
                self.add("H1.21.REFSPACING", b, f"reference spacing before={before} after={after}; single needs 0/<=12pt, 1.5 needs 0/6pt")
            if eff.get("jc") != BODY_JC:
                self.add("H1.21.REFALIGN", b, f"reference alignment {eff.get('jc')}; must match body alignment (justified)")
            if not eff.get("keepLines", False) or eff.get("widowControl") is False:
                self.add("H1.21.REFKEEP", b, "reference needs widow/orphan control and keep lines together")
            low = b.text.lower()
            if "doi" not in low and "http" not in low:
                self.add("H1.21.REFLINK", b, "reference must include a DOI or a web link")

    # ------------------------------------------------ front matter
    def rule_title_page(self):
        d = self.d
        title_blocks = [b for b in d.blocks if b.zone == "TITLE" and b.kind == "p" and b.text]
        if not title_blocks:
            self.add("H1.TITLE", "document", "title page not found in section 0")
            return
        t = title_blocks[0]
        sizes = {r.get("sz") for _, r, _ in d.runs(t.el)}
        if sizes != {28}:
            self.add("H1.TITLE", t, f"thesis title must be 14pt (found {sorted(s/2 for s in sizes if s)})")
        if t.text.upper() != t.text:
            self.add("H1.TITLE", t, "thesis title must be in ALL CAPS")
        for b in title_blocks[1:]:
            for _, r, text in d.runs(b.el):
                if r.get("sz") != 24 and re.search(r"[A-Za-z]", text):
                    self.add("H1.TITLE", b, f"title page text {text[:30]!r} is {r.get('sz', 0)/2}pt; must be 12pt")
                    break
        # Template weights: title, author, "A Thesis", degree bold; "Submitted..." and "In Partial..." italic; rest regular.
        texts = [b.text for b in title_blocks]
        degree = next((b for i, b in enumerate(title_blocks) if i > 0 and texts[i - 1].startswith("In Partial Fulfillment")), None)
        if degree is None:
            self.add("H1.TITLE", "document", "title page has no degree line after 'In Partial Fulfillment of the Requirements for the degree of'")
        want_bold = {t.text for t in title_blocks[:1]} | {texts[2] if len(texts) > 2 else "", "A Thesis", "A Dissertation"} | ({degree.text} if degree is not None else set())
        want_italic = {t for t in texts if t.startswith(("Submitted to the Faculty", "In Partial Fulfillment"))}
        for b in title_blocks:
            runs = list(d.runs(b.el))
            if not runs:
                continue
            bold = all(r.get("b") for _, r, _ in runs)
            ital = all(r.get("i") for _, r, _ in runs)
            if (b.text in want_bold) != bold:
                self.add("TPL.TITLEBOLD", b, f"title page line must be {'bold' if b.text in want_bold else 'regular weight'} as in the template (found bold={bold})")
            if (b.text in want_italic) != ital:
                self.add("TPL.TITLEBOLD", b, f"title page line must be {'italic' if b.text in want_italic else 'upright'} as in the template (found italic={ital})")
        committee = [b for b in d.blocks if b.kind == "p" and b.zone == "COMMITTEE" and b.text]
        # Template weights: both header lines, every committee-member line and "Approved by:" are bold; departments and the approver name are not.
        for b in committee:
            runs = list(d.runs(b.el))
            if not runs:
                continue
            up = b.text.upper()
            is_member = bool(re.search(r"\b(CHAIR|COMMITTEE MEMBER|COMMITTEE)\b", up)) and up not in COMMITTEE_HEADERS
            want = up in COMMITTEE_HEADERS or is_member or up.startswith("APPROVED BY")
            if any(bool(r.get("b")) != want for _, r, _ in runs):
                self.add("TPL.COMMITTEEBOLD", b, f"committee page line must be {'bold' if want else 'regular weight'} as in the template")
        if not any("COMMITTEE APPROVAL" in b.text.upper() for b in committee):
            self.add("H1.COMMITTEE", "document", "Statement of Committee Approval page not found as its own section after the title page")
        for b in committee:
            up = b.text.upper()
            want = 28 if up in COMMITTEE_HEADERS else 24
            for _, r, text in d.runs(b.el):
                if r.get("sz") != want and re.search(r"[A-Za-z]", text):
                    self.add("H1.COMMITTEE", b, f"committee page text {text[:30]!r} is {(r.get('sz') or 0)/2}pt; template uses {want/2:.0f}pt here")
                    break
            if d.eff(b).get("jc") != "center":
                self.add("H1.COMMITTEE", b, "committee page lines are centered in the template")

    def rule_toc_entries(self):
        d = self.d
        headings = []
        caps = {"Table": [], "Figure": []}
        for b in d.blocks:
            if b.pkind in ("heading1", "major") and b.zone not in ("ACKNOWLEDGMENTS", "TABLE OF CONTENTS"):
                headings.append(re.sub(r"\s+", " ", b.text))
            if b.pkind == "sub":
                headings.append(re.sub(r"\s+", " ", b.text))
            if b.pkind in ("tablecap", "figcap") and self._cap_number(b):
                caps[self._cap_number(b)[0]].append(re.sub(r"\s+", " ", b.text))
        toc = [b for b in d.blocks if b.pkind == "toc"]
        if not toc:
            self.add("H1.5.TOC", "document", "no table of contents entries found")
            return
        def entry_text(b):
            t = b.text.split("\t")[0]
            t = re.sub(r"\s+PAGEREF.*$", "", t)
            t = re.sub(r"\s*\d+\s*$", "", t) if "\t" not in b.text else t
            return re.sub(r"\s+", " ", t).strip()
        toc_entries = [entry_text(b) for b in toc if b.zone == "TABLE OF CONTENTS"]
        if any(e.upper() == "ACKNOWLEDGMENTS" for e in toc_entries):
            self.add("H1.12.NORMALHEAD", "table of contents", "ACKNOWLEDGMENTS must not appear in the table of contents")
        for b in toc:
            if b.zone == "TABLE OF CONTENTS" and b.style == "TOC1":
                e = entry_text(b)
                if e and e.upper() != e:
                    self.add("H1.5.TOCCAPS", b, "level-1 table of contents entry must be in all caps")
        strip_num = lambda s: re.sub(r"^\d+(\.\d+)*\.?\s*", "", s)
        heading_set = {strip_num(h) for h in headings}
        for b in toc:
            if b.zone != "TABLE OF CONTENTS":
                continue
            e = strip_num(entry_text(b))
            if e and e not in heading_set and e != "TABLE OF CONTENTS":
                self.add("H1.5.TOCSTALE", b, f"table of contents entry {e!r} does not match any heading; update the field")
        for label, zone in (("Table", "LIST OF TABLES"), ("Figure", "LIST OF FIGURES")):
            entries = [entry_text(b) for b in toc if b.zone == zone and entry_text(b)]
            want = [re.sub(r"\s+", " ", c) for c in caps[label]]
            if entries and entries != want:
                missing = [w for w in want if w not in entries]
                extra = [e for e in entries if e not in want]
                if missing or extra or len(entries) != len(want):
                    self.add("H1.5.TOCSTALE", zone, f"{zone} is out of date: {len(missing)} caption(s) missing, {len(extra)} stale entr(ies); update the field. missing={missing[:2]} extra={extra[:2]}")
        for b in toc:
            eff = d.eff(b)
            for _, r, text in d.runs(b.el):
                if r.get("sz") != 24 and re.search(r"[A-Za-z]", text):
                    self.add("H1.8.SIZE", b, f"list entry {text[:30]!r} is {r.get('sz', 0)/2}pt; must be 12pt")
                    break

    def rule_extra_headings_sorted(self):
        """Handbook 'EXTRA HEADINGS' page: glossary / abbreviation / symbol entries sorted alphabetically."""
        d = self.d
        for zone in NO_INDENT_ZONES:
            keys = [b.text.split("\t")[0].strip() for b in d.blocks if b.zone == zone and b.pkind == "body" and "\t" in b.text]
            for a, c in zip(keys, keys[1:]):
                if a.lower() > c.lower():
                    self.add("H1.EXTRA.SORTED", zone, f"entries {a!r} and {c!r} are out of alphabetical order")

    def rule_revisions(self):
        """A deposited document carries no tracked changes or comments."""
        d = self.d
        marks = Counter(local(e) for e in d.doc.iter() if local(e) in REVISION_TAGS)
        if marks:
            self.add("REV.TRACKED", "document", f"tracked changes remain ({dict(marks)}); accept or reject every change before submission")
        if d.settings is not None and d.settings.find("w:trackRevisions", NS) is not None:
            self.add("REV.TRACKED", "settings", "Track Changes is switched on; turn it off so the file reviewers open is final")
        if d.has_comments and d.doc.find(".//w:commentRangeStart", NS) is not None:
            self.add("REV.TRACKED", "document", "comments remain in the document; delete them before submission")

    def rule_abstract(self):
        d = self.d
        paras = [b for b in d.blocks if b.zone == "ABSTRACT" and b.pkind == "body"]
        if len(paras) > 1:
            self.add("H1.ABSTRACT", paras[1], f"abstract has {len(paras)} paragraphs; the handbook asks for a single un-indented paragraph", "INFO")
        for b in paras:
            if (d.eff(b).get("firstLine") or 0) != 0:
                self.add("H1.ABSTRACT", b, "abstract paragraph should not be indented", "INFO")
            words = len(b.text.split())
            if words > ABSTRACT_LIMIT:
                self.add("H1.ABSTRACT", b, f"abstract is {words} words; aim for {ABSTRACT_LIMIT} or fewer", "INFO")


# ---------------------------------------------------------------- reporting
def report(issues: list[Issue], title: str, limit: int = 12) -> int:
    errors = [i for i in issues if i.severity == "ERROR"]
    infos = [i for i in issues if i.severity != "ERROR"]
    print(f"=== {title}")
    by_rule: dict[str, list[Issue]] = defaultdict(list)
    for i in issues:
        by_rule[i.rule].append(i)
    for rule in sorted(by_rule):
        items = by_rule[rule]
        sev = items[0].severity
        print(f"[{sev}] {rule}: {len(items)}")
        for it in items[:limit]:
            print(f"    - {it.loc}: {it.msg}")
        if len(items) > limit:
            print(f"    ... and {len(items) - limit} more")
    print(f"--- {len(errors)} error(s), {len(infos)} info note(s) -> {'PASS' if not errors else 'FAIL'}")
    return len(errors)


def check(path: str) -> list[Issue]:
    return Checker(Document(path)).run()


# ---------------------------------------------------------------- self-test
def _mutate_and_check(doc: Document, mutate) -> Counter:
    """Deep-copy the document XML, apply `mutate(doc_copy)`, return rule counts."""
    d2 = copy.copy(doc)
    d2.doc = copy.deepcopy(doc.doc)
    d2.body = d2.doc.find("w:body", NS)
    d2.styles = copy.deepcopy(doc.styles)
    d2.numbering = Numbering(etree.tostring(doc.numbering.root) if doc.numbering.root is not None else None)
    d2.settings = copy.deepcopy(doc.settings)
    d2.footnotes = copy.deepcopy(doc.footnotes)
    d2.parts = {k: copy.deepcopy(v) for k, v in doc.parts.items()}
    d2.media = dict(doc.media)
    d2._palette_cache = {}
    d2.blocks, d2.sects = [], []
    d2._walk()
    mutate(d2)
    d2.styles.refresh()
    d2.blocks, d2.sects = [], []
    d2._walk()
    return Counter(i.rule for i in Checker(d2).run())


def _set_ppr(p, tag, **attrs):
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.SubElement(p, q("pPr"))
        p.insert(0, ppr)
    el = ppr.find(f"w:{tag}", NS)
    if el is None:
        el = etree.SubElement(ppr, q(tag))
    for k, v in attrs.items():
        el.set(q(k), str(v))
    return el


def _first(doc: Document, pkind, zone_prefix=None):
    for b in doc.blocks:
        if b.pkind == pkind and (zone_prefix is None or b.zone.startswith(zone_prefix)):
            return b
    raise LookupError(pkind)


def _first_tbl(doc: Document, with_img=False):
    for b in doc.blocks:
        if b.kind == "tbl" and ((b.el.find(".//w:drawing", NS) is not None) == with_img):
            return b
    raise LookupError("tbl")


def _fitting_tbl(doc: Document):
    for b in doc.blocks:
        if b.kind == "tbl" and b.el.find(".//w:drawing", NS) is None and doc.table_height_pt(b.el) <= TABLE_FITS_PT:
            return b
    raise LookupError("fitting tbl")


def _drop_from(b: Block, path: str):
    el = b.el.find(path, NS)
    el.getparent().remove(el)


def _noted_tbl(doc: Document):
    for b in doc.blocks:
        nxt = doc.next(b)
        if b.kind == "tbl" and nxt is not None and nxt.pkind == "note":
            return b
    raise LookupError("tbl followed by note")


def _add_comment(d):
    d.has_comments = True
    etree.SubElement(_first(d, "body", "CH").el, q("commentRangeStart")).set(q("id"), "0")


def _fake_heading(d):
    b = _first(d, "sub")
    b.el.find("w:pPr/w:pStyle", NS).set(q("val"), "Normal")
    _set_run_rpr(b.el, "b")


def _duplicate_alt(d):
    figs = [b for b in d.blocks if b.pkind == "figure" and b.zone.startswith("CH")][:2]
    for b in figs:
        next(b.el.iter(f"{{{WP}}}docPr")).set("descr", "Identical text on two different pictures")


def _low_contrast_picture(d):
    import io
    from PIL import Image, ImageDraw
    b = _first(d, "figure", "CH")
    media = d.drawing_media(next(b.el.iter(q("drawing"))))
    im = Image.new("RGB", (200, 120), "white")
    # #FFE566 bar (1.3:1 on white), ~8% of the picture so it is a chart element, not an area fill
    ImageDraw.Draw(im).rectangle([20, 60, 60, 108], fill=(255, 229, 102))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    d.media[media] = buf.getvalue()


def _page_sized_picture(d):
    b = _first(d, "figure", "CH")
    ext = next(b.el.iter(f"{{{WP}}}extent"))
    ext.set("cx", str(int(6.5 * 914400))); ext.set("cy", str(int(9 * 914400)))
    cap = d.next(b).el
    _retext(cap, "No caption follows this picture any more.")
    for ps in cap.findall("w:pPr/w:pStyle", NS):
        ps.getparent().remove(ps)


def _wrap_in_layout_table(d):
    b = _first(d, "figure", "CH")
    tbl = etree.Element(q("tbl"))
    etree.SubElement(tbl, q("tblPr"))
    tr = etree.SubElement(tbl, q("tr"))
    tc = etree.SubElement(tr, q("tc"))
    b.el.addprevious(tbl)
    tc.append(b.el)
    tr2 = etree.SubElement(tbl, q("tr"))
    tc2 = etree.SubElement(tr2, q("tc"))
    p = etree.SubElement(tc2, q("p"))
    _retext(p, "(a)")


def _ordered_tcpr(tc):
    pr = tc.find("w:tcPr", NS)
    if pr is None:
        pr = etree.Element(q("tcPr"))
        tc.insert(0, pr)
    return pr


def _text_box_with_text(d):
    b = _first(d, "body", "CH")
    r = etree.SubElement(b.el, q("r"))
    tx = etree.SubElement(r, q("txbxContent"))
    _retext(etree.SubElement(tx, q("p")), "Text hidden in a box")


def _unkeep_last_row(b: Block):
    for p in b.el.findall("w:tr", NS)[-1].iter(q("p")):
        kn = p.find("w:pPr/w:keepNext", NS)
        if kn is not None:
            kn.getparent().remove(kn)


def _set_run_rpr(p, tag, **attrs):
    for r in p.iter(q("r")):
        if r.find("w:t", NS) is None:
            continue
        rpr = r.find("w:rPr", NS)
        if rpr is None:
            rpr = etree.Element(q("rPr"))
            r.insert(0, rpr)
        el = rpr.find(f"w:{tag}", NS)
        if el is None:
            el = etree.SubElement(rpr, q(tag))
        for k, v in attrs.items():
            el.set(q(k), str(v))


def selftest(path: str) -> bool:
    """Four review passes. Each pass injects one known violation into a copy of
    the document and asserts that the matching rule fires (its count rises
    above the baseline). Run it against a document that already passes so the
    baseline is zero and every rule is proven from a clean state."""
    doc = Document(path)
    base = Counter(i.rule for i in Checker(doc).run())

    def body_p(d):
        return _first(d, "body", "CH").el

    def fig(d):
        return _first(d, "figure", "CH").el

    passes = {
        "PASS 1 - spaces, gaps, line spacing, indents": [
            ("H1.9.SPACING", lambda d: _set_ppr(body_p(d), "spacing", line=240, lineRule="auto")),
            ("H1.10.INDENT", lambda d: _set_ppr(body_p(d), "ind", firstLine=360)),
            ("H1.10.PARASPACE", lambda d: _set_ppr(body_p(d), "spacing", after=200)),
            ("H1.10.EMPTYPARA", lambda d: body_p(d).addnext(etree.Element(q("p")))),
            ("H1.11.ALIGN", lambda d: _set_ppr(body_p(d), "jc", val="left")),
            ("H1.12.NOSPACE", lambda d: _first(d, "sub").el.addnext(etree.Element(q("p")))),
            ("H1.12.SUBFMT", lambda d: _set_ppr(_first(d, "sub").el, "spacing", before=0)),
            ("H1.12.HEADFMT", lambda d: _set_ppr(_first(d, "heading1").el, "spacing", after=240)),
            ("H1.13.CAPSPACE", lambda d: _set_ppr(_first(d, "tablecap").el, "spacing", before=0)),
            ("H1.13.CAPKEEP", lambda d: _set_ppr(_first(d, "figcap").el, "keepLines", val="0")),
            ("H1.13.FIGGAP", lambda d: _set_ppr(fig(d), "spacing", before=0)),
            ("H1.13.TBLAFTER", lambda d: _set_ppr(d.next(_first_tbl(d)).el, "spacing", before=0)),
            ("H1.13.TBLSPACING", lambda d: _set_ppr(_first_tbl(d).el.find(".//w:p", NS), "spacing", line=360, lineRule="auto")),
            ("H1.9.FRONT", lambda d: _set_ppr([b for b in d.blocks if b.zone == "COMMITTEE" and b.text][0].el, "spacing", line=240, lineRule="auto")),
            ("H1.4.MARGIN", lambda d: d.sects[-1].find("w:pgMar", NS).set(q("left"), "1800")),
            ("H1.21.REFHANG", lambda d: _set_ppr(_first(d, "ref").el, "ind", hanging=0)),
            ("H1.21.REFSPACING", lambda d: _set_ppr(_first(d, "ref").el, "spacing", before=200)),
        ],
        "PASS 2 - fonts, font sizes, text, headings case": [
            ("H1.8.FONT", lambda d: _set_run_rpr(body_p(d), "rFonts", ascii="Arial", hAnsi="Arial")),
            ("H1.8.SIZE", lambda d: _set_run_rpr(body_p(d), "sz", val=22)),
            ("TPL.COLOR", lambda d: _set_run_rpr(body_p(d), "color", val="FF0000")),
            ("H1.12.CAPS", lambda d: _lower(_first(d, "heading1").el)),
            ("H1.12.SENTENCECASE", lambda d: _upper(_first(d, "sub").el)),
            ("H1.12.STYLE", lambda d: _styled_major(d).find("w:pPr/w:pStyle", NS).set(q("val"), "Normal")),
            ("H1.12.NUMBER", lambda d: _set_ppr(_first(d, "heading1").el, "numPr").append(_num(0))),
            ("H1.13.TABLETEXT", lambda d: _set_run_rpr(_first_tbl(d).el.find(".//w:p", NS), "sz", val=16)),
            ("H1.13.NOTE", lambda d: _add_note(d)),
            ("H1.TITLE", lambda d: _set_run_rpr([b for b in d.blocks if b.zone == "TITLE" and b.text][0].el, "sz", val=24)),
            ("H1.COMMITTEE", lambda d: _set_run_rpr([b for b in d.blocks if b.zone == "COMMITTEE" and b.text][2].el, "sz", val=28)),
            ("H1.18.FOOTNOTE", lambda d: _add_footnote(d)),
            ("H1.21.REFNUM", lambda d: _retext(_first(d, "ref").el, "[99] Broken reference. https://x")),
            ("H1.21.REFLINK", lambda d: _retext(_first(d, "ref").el, "[1] Reference without a link.")),
            ("H1.7.PAGENUM", lambda d: _strip_footer_page(d)),
        ],
        "PASS 3 - images, tables, columns, captions": [
            ("H1.13.FIGCENTER", lambda d: _set_ppr(fig(d), "jc", val="left")),
            ("H1.13.FIGSPACING", lambda d: _set_ppr(fig(d), "spacing", line=360, lineRule="auto")),
            ("H1.13.SAMEPAGE", lambda d: _set_ppr(fig(d), "keepNext", val="0")),
            ("H1.13.ALTTEXT", lambda d: next(fig(d).iter(f"{{{WP}}}docPr")).set("descr", "")),
            ("H1.13.WRAP", lambda d: _anchor_drawing(fig(d))),
            ("H1.13.TBLCENTER", lambda d: _set_tbl_jc(_first_tbl(d).el, "left")),
            ("H1.13.CAPORDER", lambda d: _first(d, "tablecap").el.addnext(etree.Element(q("p")))),
            ("H1.13.CAPSTYLE", lambda d: _first(d, "figcap").el.find("w:pPr/w:pStyle", NS).set(q("val"), "Normal")),
            ("H1.13.CAPNUM", lambda d: _retext(_first(d, "figcap").el, "Figure 9.9. Wrong chapter number.")),
            ("H1.13.INSERTCAP", lambda d: _retext(_first(d, "figcap").el, "Figure 1.1. Typed by hand.")),
            ("SK.CAPALIGN", lambda d: _set_ppr(_first(d, "figcap").el, "jc", val="left")),
            ("SK.XREFORDER", lambda d: _xref_after(d)),
            ("H1.13.CAPINTABLE", lambda d: _caption_into_cell(d)),
            ("H1.13.INLINEPARA", lambda d: _retext(fig(d), "stray text next to picture", keep=True)),
            ("H1.15.XREF", lambda d: _retext(_first(d, "figcap").el, "Figure 1.9. Never referenced anywhere at all.")),
            ("H1.13.NESTED", lambda d: _first_tbl(d).el.find(".//w:tc", NS).append(etree.Element(q("tbl")))),
        ],
        "PASS 4 - headings, sections, breaks, lists, references, equations": [
            ("H1.6.PAGEBREAK", lambda d: _remove_pagebreak_before(d, _first(d, "heading1"))),
            ("H1.6.CONTINUOUS", lambda d: etree.SubElement(d.sects[0], q("type")).set(q("val"), "continuous")),
            ("H1.7.TITLEPAGE", lambda d: _add_page_field_to_title_footer(d)),
            ("H1.5.TOCSTALE", lambda d: _retext([b for b in d.blocks if b.pkind == "toc" and b.zone == "TABLE OF CONTENTS" and b.style == "TOC1"][1].el, "A HEADING THAT DOES NOT EXIST")),
            ("H1.5.TOCCAPS", lambda d: _retext([b for b in d.blocks if b.pkind == "toc" and b.zone == "TABLE OF CONTENTS" and b.style == "TOC1"][1].el, "Lower case entry")),
            ("H1.12.NORMALHEAD", lambda d: _retext([b for b in d.blocks if b.pkind == "toc" and b.zone == "TABLE OF CONTENTS" and b.style == "TOC1"][1].el, "ACKNOWLEDGMENTS")),
            ("H1.DEDICATION", lambda d: _set_ppr([b for b in d.blocks if b.zone == "DEDICATION" and b.text][0].el, "jc", val="left")),
            ("H1.16.EQNUM", lambda d: _add_equation(d)),
            ("H1.17.BLOCKQUOTE", lambda d: _first(d, "body", "CH").el.find("w:pPr", NS).insert(0, _pstyle("BlockQuote"))),
            ("H1.21.REFALIGN", lambda d: _set_ppr(_first(d, "ref").el, "jc", val="center")),
            ("H1.21.REFKEEP", lambda d: _set_ppr(_first(d, "ref").el, "keepLines", val="0")),
            ("H1.12.NUMBER", lambda d: _set_ppr(_first(d, "sub").el, "numPr").append(_num(0))),
            ("H1.18.FOOTNOTE", lambda d: etree.SubElement(_first(d, "sub").el, q("r")).append(etree.Element(q("footnoteReference")))),
        ],
        "PASS 5 - template numbering, blank pages, table integrity, front-page weights, revisions, sorting": [
            ("H1.12.NUMIND", lambda d: _heading_lvl(d, "Heading3").find("w:pPr/w:ind", NS).set(q("left"), "4410")),
            ("H1.12.NUMFMT", lambda d: _heading_lvl(d, "Heading3").find("w:lvlText", NS).set(q("val"), "%1-%2-%3")),
            ("H1.12.SUBFMT", lambda d: _set_ppr(_sub_at_level(d, 3).el, "jc", val="both")),
            ("H1.6.BLANKPAGE", lambda d: _pagebreak_paragraph_before(d, _first(d, "heading1"))),
            ("H1.13.TBLKEEP", lambda d: _drop_from(_fitting_tbl(d), "w:tr/w:trPr/w:cantSplit")),
            ("H1.13.TBLHEADER", lambda d: _first_tbl(d).el.find("w:tr/w:trPr/w:tblHeader", NS).getparent().remove(_first_tbl(d).el.find("w:tr/w:trPr/w:tblHeader", NS))),
            ("H1.13.TBLFONTMIX", lambda d: _set_run_rpr(_first_tbl(d).el.find(".//w:p", NS), "sz", val=20)),
            ("H1.13.TBLHEADFMT", lambda d: _toggle_header_bold(d)),
            ("H1.13.TBLLONG", lambda d: _lengthen_table(d)),
            ("H1.13.TBLNOTE", lambda d: _unkeep_last_row(_noted_tbl(d))),
            ("H1.13.TBLJC", lambda d: _set_ppr(_first_tbl(d).el.find(".//w:p", NS), "jc", val="both")),
            ("TPL.TITLEBOLD", lambda d: _set_run_rpr([b for b in d.blocks if b.zone == "TITLE" and b.text][2].el, "b", val="0")),
            ("TPL.COMMITTEEBOLD", lambda d: _set_run_rpr([b for b in d.blocks if b.zone == "COMMITTEE" and b.text][3].el, "b")),
            ("REV.TRACKED", lambda d: _add_comment(d)),  # a comment is the one revision kind a reviewer copy never carries
            ("H1.EXTRA.SORTED", lambda d: _swap_abbreviations(d)),
        ],
        "PASS 6 - accessibility (WCAG 2.1 AA)": [
            ("A11Y.HEADING", lambda d: _fake_heading(d)),
            ("A11Y.ALT", lambda d: next(fig(d).iter(f"{{{WP}}}docPr")).set("descr", "A picture containing text. Description automatically generated")),
            ("A11Y.ALTDUP", lambda d: _duplicate_alt(d)),
            ("A11Y.CONTRAST", lambda d: _low_contrast_picture(d)),
            ("A11Y.SCANNED", lambda d: _page_sized_picture(d)),
            ("A11Y.LAYOUTTBL", lambda d: _wrap_in_layout_table(d)),
            ("A11Y.TBLMERGE", lambda d: etree.SubElement(_ordered_tcpr(_first_tbl(d).el.find(".//w:tc", NS)), q("gridSpan")).set(q("val"), "2")),
            ("A11Y.TEXTBOX", lambda d: _text_box_with_text(d)),
            ("H1.13.WRAP", lambda d: _anchor_drawing([b for b in d.blocks if b.zone == "TITLE" and b.has_drawing][0].el)),
        ],
    }
    all_ok = True
    for title, tests in passes.items():
        print(f"=== SELFTEST {title}")
        for rule, mut in tests:
            try:
                counts = _mutate_and_check(doc, mut)
                fired = counts[rule] > base[rule]
            except Exception as exc:  # a mutation helper could not find its target
                fired = False
                counts = Counter({"exception": 1})
                print(f"    ! {rule}: mutation raised {exc!r}")
            print(f"    {'OK  ' if fired else 'MISS'} {rule}  (baseline {base[rule]} -> mutated {counts[rule]})")
            all_ok &= fired
        print()
    print("SELFTEST RESULT:", "ALL RULES FIRE" if all_ok else "SOME RULES DID NOT FIRE")
    return all_ok


# -- mutation helpers
def _retext(p, text, keep=False):
    if not keep:
        for c in list(p):
            if local(c) not in ("pPr", "bookmarkStart", "bookmarkEnd"):
                p.remove(c)
    r = etree.SubElement(p, q("r"))
    t = etree.SubElement(r, q("t"))
    t.text = text


def _styled_major(d):
    for b in d.blocks:
        if b.pkind == "major" and b.el.find("w:pPr/w:pStyle", NS) is not None and b.text.upper() not in NORMAL_HEADINGS:
            return b.el
    raise LookupError("styled major heading")


def _anchor_drawing(p):
    inl = next(p.iter(f"{{{WP}}}inline"))
    inl.tag = f"{{{WP}}}anchor"


def _set_tbl_jc(tbl, val):
    tblpr = tbl.find("w:tblPr", NS)
    jc = tblpr.find("w:jc", NS)
    if jc is None:
        jc = etree.SubElement(tblpr, q("jc"))
    jc.set(q("val"), val)


def _xref_after(d):
    cap = _first(d, "figcap")
    _retext(cap.el, "Figure 1.9. Mentioned only after it appears.")
    later = next(b for b in d.blocks if b.idx > cap.idx and b.pkind == "body")
    _retext(later.el, " See Figure 1.9 for the layout.", keep=True)


def _lower(p):
    for t in p.iter(q("t")):
        t.text = (t.text or "").lower()


def _upper(p):
    for t in p.iter(q("t")):
        t.text = (t.text or "").upper()


def _num(v):
    el = etree.Element(q("numId"))
    el.set(q("val"), str(v))
    return el


def _pstyle(v):
    el = etree.Element(q("pStyle"))
    el.set(q("val"), v)
    return el


def _add_note(d):
    tbl = _first_tbl(d).el
    p = etree.Element(q("p"))
    _retext(p, "Note: this is a note without italics.")
    tbl.addnext(p)


def _add_footnote(d):
    fn = etree.SubElement(d.footnotes, q("footnote"))
    fn.set(q("id"), "99")
    p = etree.SubElement(fn, q("p"))
    _set_ppr(p, "spacing", line=360, lineRule="auto")
    _retext(p, "Footnote with wrong spacing")


def _strip_footer_page(d):
    for part in d.parts.values():
        for t in part.iter(q("instrText")):
            if "PAGE" in (t.text or ""):
                t.text = "NOTHING"


def _caption_into_cell(d):
    tc = _first_tbl(d).el.find(".//w:tc", NS)
    p = etree.SubElement(tc, q("p"))
    _set_ppr(p, "pStyle", val="TableCaption")
    _retext(p, "Table 2.1. Caption inside a cell.")


def _remove_pagebreak_before(d, heading):
    prev = d.prev(heading)
    for br in list(prev.el.iter(q("br"))):
        br.getparent().remove(br)
    for sp in list(prev.el.iter(q("sectPr"))):
        sp.getparent().remove(sp)
    for pb in heading.el.findall("w:pPr/w:pageBreakBefore", NS):
        pb.getparent().remove(pb)
    _retext(prev.el, "Text so the heading no longer starts a page.")


def _toggle_header_bold(d):
    row0 = _first_tbl(d).el.find("w:tr", NS)
    p = row0.find(".//w:p", NS)
    _, eff, _ = next(d.runs(p))
    _set_run_rpr(p, "b", val="0" if eff.get("b") else "1")


def _add_page_field_to_title_footer(d):
    sp = d.sects[0]
    for r in sp:
        if local(r) == "footerReference" and r.get(q("type")) == "first":
            part = d.parts[r.get(f"{{{R}}}id")]
            p = part.find(".//w:p", NS)
            run = etree.SubElement(p, q("r"))
            it = etree.SubElement(run, q("instrText"))
            it.text = " PAGE "


def _add_equation(d):
    p = _first(d, "body", "CH").el
    om = etree.SubElement(p, f"{{{M}}}oMath")
    etree.SubElement(om, f"{{{M}}}r")


def _heading_lvl(d, style):
    b = next(b for b in d.blocks if b.style == style)
    eff = d.eff(b)
    return d.numbering.level(eff["numId"], eff.get("ilvl", 0))


def _sub_at_level(d, lvl):
    return next(b for b in d.blocks if b.pkind == "sub" and SUB_STYLES[b.style] >= lvl)


def _pagebreak_paragraph_before(d, heading):
    ppr = heading.el.find("w:pPr", NS)
    for pb in ppr.findall("w:pageBreakBefore", NS):
        ppr.remove(pb)
    p = etree.Element(q("p"))
    etree.SubElement(etree.SubElement(p, q("r")), q("br")).set(q("type"), "page")
    heading.el.addprevious(p)


def _lengthen_table(d):
    tbl = _fitting_tbl(d).el
    last = tbl.findall("w:tr", NS)[-1]
    for _ in range(60):
        last.addnext(copy.deepcopy(last))


def _swap_abbreviations(d):
    entries = [b for b in d.blocks if b.zone in NO_INDENT_ZONES and b.pkind == "body" and "\t" in b.text]
    a, c = entries[0].el, entries[1].el
    ta, tc = a.find(".//w:t", NS), c.find(".//w:t", NS)
    ta.text, tc.text = tc.text, ta.text


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    path = argv[1]
    for i, a in enumerate(argv[2:], 2):
        if a == "--university" and i + 1 < len(argv):
            apply_config(load_university(argv[i + 1]))
        elif a == "--config" and i + 1 < len(argv):
            apply_config(load_university(argv[i + 1]))
    if "--selftest" in argv:
        return 0 if selftest(path) else 1
    issues = check(path)
    errors = report(issues, f"FORMAT CHECK {path.rsplit('/', 1)[-1]}")
    return 0 if errors == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
