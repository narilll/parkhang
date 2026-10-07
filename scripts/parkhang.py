#!/usr/bin/env python3
"""Typeset an interlinear text.md (lotsawa output) into a Word .docx in the house print style.

Usage:
  parkhang.py build text.md -o out.docx [--dry-run] [--header1 S] [--header2 S] [--footer S]
  parkhang.py selftest

Segments are classified heuristically (see SKILL.md); optional extensions: front matter
(header1/header2/footer/cover_*/tib_font/notes_label; header2 defaults to the title), ` >> CUE`, `{SYL}`, `*italic*`,
`#`/`##` heading override, `—` / 2+ blank lines as separator, endnote markers (`word,ii`,
`STEZKYi`, `slovo^v`) matched to the lines under the `Poznámky:` label.
Data loss (marker/note count mismatch, missing cover image) is fatal (exit 1); an
editorial judgment call is a warning (exit 0). --dry-run lists paragraph styles only.
Stdlib only; raw OOXML via zipfile.
"""
import argparse
import contextlib
import io
import re
import struct
import sys
import tempfile
import zipfile
from collections import namedtuple
from itertools import groupby
from pathlib import Path
from xml.sax.saxutils import escape

TIB = 'ༀ-࿿'
TIB_RE = re.compile(f'[{TIB}]')
TIB_SPLIT = re.compile(f'([{TIB}]+)')
RED, BLUE = 'C00000', '007BB8'
TAB_POS = 9071                       # right tab = 16 cm
NOTE_PUNCT = ',.;:!?…)]"“”»’'
INLINE_RE = re.compile(r'\{([^}]*)\}|\*([^*]+)\*|\^?([ivxlc]+)(?=[\s,.;:!?)\]"”»]|$)')
LINK_RE = re.compile(r'\[([^\]]*)\]\(([^)\s]+)\)')
CUE_RE = re.compile(r'^(.*?)\s*>>\s*(.+)$')
H1_TIB_RE = re.compile('ནི\\s*[།༔][།༔\\s]*$')
VERSE_END_RE = re.compile('(།\\s*།|༔)\\s*$')
# ponytail: prose instructions end with the terminative -o particle (བྱའོ། །, བཟུང་ངོ་། །); verses rarely do
PROSE_END_RE = re.compile('ོ་?[།\\s]+$')
MANTRA_IN_RE = re.compile(r'\b[A-ZÄÖÜČŠŽŇ]{2,}\s+[A-ZÄÖÜČŠŽŇ]{2,}\b')   # mantra words inside an instruction
IMPERATIVE_RE = re.compile(r'\b\w+(?:te|ete|ěte|jte)\b')
COVER_BOX = (3564000, 5292000)       # 99 x 147 mm in EMU
TIB_FONT = 'Jomolhari'               # overridden by front-matter tib_font
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = R_NS
NS = f'xmlns:w="{W}" xmlns:r="{R_NS}"'
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
# ponytail: framePr yAlign=bottom lands on the next page in LibreOffice; fixed 36 pt gap above the
# cover bottom block instead. Pin to the bottom margin by hand in Word if needed.
GAP = '<w:spacing w:before="720"/>'

_warn = 0
_err = 0


def warn(sub, msg):
    global _warn
    _warn += 1
    print(f'WARN {sub}: {msg}', file=sys.stderr)


def err(sub, msg):
    global _err
    _err += 1
    print(f'ERROR {sub}: {msg}', file=sys.stderr)


def info(msg):
    print(f'INFO {msg}', file=sys.stderr)


def finish(sub, exit_on_error=True):
    print(f'{_warn} warnings, {_err} errors', file=sys.stderr)
    if _err and exit_on_error:
        sys.exit(1)
    return _err


# ---------------------------------------------------------------- classification

def has_tib(s):
    return bool(TIB_RE.search(s))


def looks_like_pho(line):
    """Phonetics line: uniform casing and no sentence-level punctuation (copy from lotsawa)."""
    s = line.strip()
    if not s:
        return False
    if re.search(r'[,!?;:]', s):
        return False
    letters = [ch for ch in s if ch.isalpha()]
    if not letters:
        return False
    upper = sum(1 for ch in letters if ch.isupper())
    return upper == len(letters) or upper == 0


def demark(s):
    """Drops {} * and endnote markers (`STEZKYi`, `^v`) so case checks see the bare text."""
    s = re.sub(r'(?<=\S)\^[ivxlc]+', '', re.sub(r'[{}*]', '', s))
    return re.sub(r'([^\W\d_])([ivxlc]+)(?=[\s,.;:!?)\]"”»]|$)',
                  lambda m: m.group(1) if m.group(1).isupper() else m.group(0), s)


def is_pho(s):
    return looks_like_pho(demark(s))


def is_caps(s):
    letters = [c for c in demark(s) if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters)


def short(s):
    """Heading-like: <= 8 words, first letter not lowercase, no closing punctuation, no Czech 2nd-pl imperative, not all italic."""
    # ponytail: Czech-only heuristic; every use is logged as INFO for the agent to review
    s = s.strip()
    if s.startswith('*') and s.endswith('*'):
        return False
    if next((c for c in s if c.isalpha()), 'A').islower():
        return False                      # lowercase start = a continuation, never a heading
    return len(s.split()) <= 8 and s[-1:] not in '.:!,;?' and not IMPERATIVE_RE.search(s)


def roman(n):
    out = ''
    for v, s in ((10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')):
        while n >= v:
            out += s
            n -= v
    return out


def split_cue(t):
    m = CUE_RE.match(t)
    return (m.group(1).strip(), m.group(2).strip()) if m else (t.strip(), None)


# ---------------------------------------------------------------- runs and paragraphs

R = namedtuple('R', 'text b i noi rstyle color sz tib raw',
               defaults=(False, False, False, None, None, None, False, None))
P = namedtuple('P', 'lineno style runs pbb gap', defaults=(False, False))


def rpr(rstyle=None, tib=False, b=False, i=False, noi=False, caps=False, color=None, sz=None,
        sup=False):
    """rPr in schema order: rStyle, rFonts, b, i, caps, color, sz, vertAlign."""
    x = f'<w:rStyle w:val="{rstyle}"/>' if rstyle else ''
    if tib:
        f = escape(TIB_FONT, {'"': '&quot;'})
        x += f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
    if b:
        x += '<w:b/><w:bCs/>'
    if i:
        x += '<w:i/><w:iCs/>'
    elif noi:
        x += '<w:i w:val="0"/><w:iCs w:val="0"/>'
    if caps:
        x += '<w:caps/>'
    if color:
        x += f'<w:color w:val="{color}"/>'
    if sz:
        x += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'   # szCs == sz for Tibetan
    if sup:
        x += '<w:vertAlign w:val="superscript"/>'
    return f'<w:rPr>{x}</w:rPr>' if x else ''


def ppr(style=None, keep=None, pbb=False, gap=False, tabs=False, before=None, after=None,
        jc=None, outline=None):
    """pPr in schema order: pStyle, keepNext, pageBreakBefore, (gap spacing), tabs, spacing, jc, outlineLvl."""
    x = f'<w:pStyle w:val="{style}"/>' if style else ''
    if keep is not None:
        x += '<w:keepNext/>' if keep else '<w:keepNext w:val="0"/>'
    if pbb:
        x += '<w:pageBreakBefore/>'
    if gap:
        x += GAP
    if tabs:
        x += f'<w:tabs><w:tab w:val="right" w:pos="{TAB_POS}"/></w:tabs>'
    if before is not None or after is not None:
        x += '<w:spacing'
        x += f' w:before="{before}"' if before is not None else ''
        x += f' w:after="{after}"' if after is not None else ''
        x += '/>'
    if jc:
        x += f'<w:jc w:val="{jc}"/>'
    if outline is not None:
        x += f'<w:outlineLvl w:val="{outline}"/>'
    return f'<w:pPr>{x}</w:pPr>' if x else ''


def run_xml(r):
    if r.raw:
        return r.raw
    t = f'<w:t xml:space="preserve">{escape(r.text)}</w:t>'
    return f'<w:r>{rpr(r.rstyle, r.tib, r.b, r.i, r.noi, color=r.color, sz=r.sz)}{t}</w:r>'


def runs_xml(runs):
    return ''.join(run_xml(r) for r in runs)


def runs_text(runs):
    return ''.join(r.text for r in runs)


def p_xml(p):
    return f'<w:p>{ppr(p.style, pbb=p.pbb, gap=p.gap)}{runs_xml(p.runs)}</w:p>'


def plain_runs(text, **kw):
    """Plain segment; Tibetan pieces (stray Tibetan, ornaments) get the Tibetan font."""
    return [R(p, tib=bool(TIB_RE.match(p)), **kw) for p in TIB_SPLIT.split(text) if p]


def rubric_runs(text, **kw):
    """UPPERCASE tokens (>= 2 letters, e.g. OM AH HUNG) -> bold, non-italic."""
    toks = re.findall(r'\s+|\S+', text)
    flag = [not t.isspace() and t.isupper() and sum(c.isalpha() for c in t) >= 2 for t in toks]
    for k, t in enumerate(toks):             # spaces between two bold tokens stay in the bold run
        if t.isspace() and 0 < k < len(toks) - 1:
            flag[k] = flag[k - 1] and flag[k + 1]
    out = []
    for f, grp in groupby(zip(toks, flag), key=lambda x: x[1]):
        seg = ''.join(t for t, _ in grp)
        out += plain_runs(seg, b=True, noi=True, **kw) if f else plain_runs(seg, **kw)
    return out


def link_text(m):
    """`[text](url)` -> `text` if one contains the other (ignoring scheme/case), else `text (url)`."""
    t, u = m.groups()
    a, b = (re.sub(r'^https?://', '', x.strip(), flags=re.I).lower() for x in (t, u))
    return t if a in b or b in a else f'{t} ({u})'


def latin_runs(text, marks=None, ln=0, rubric=False):
    """Inline markup -> runs. `marks` (list) collects accepted endnote markers; None = no markers."""
    out, pos = [], 0
    text = LINK_RE.sub(link_text, text)
    seg = rubric_runs if rubric else plain_runs
    for m in INLINE_RE.finditer(text):
        drum, ital, num = m.groups()
        if num is not None:
            prev = text[m.start() - 1] if m.start() else ''
            caret = m.group(0).startswith('^')
            ok = caret or (prev != '' and (prev in NOTE_PUNCT or prev.isupper()))
            if marks is None or not ok or num != roman(len(marks) + 1):
                continue                      # literal text
            out += seg(text[pos:m.start()])
            marks.append((ln, m.group(0)))
            out.append(R(('^' if caret else '') + num, raw='<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
                                        f'<w:endnoteReference w:id="{len(marks)}"/></w:r>'))
        elif drum is not None:
            out += seg(text[pos:m.start()])
            out += plain_runs(drum, rstyle='Drum')
        else:
            out += seg(text[pos:m.start()])
            out += plain_runs(ital, i=True)
        pos = m.end()
    return out + seg(text[pos:])


def cue_runs(cue):
    """Tab to the right margin, then the cue; Tibetan ornaments in 10 pt red Jomolhari."""
    out = [R(' >> ', raw='<w:r><w:tab/></w:r>')]
    for p in TIB_SPLIT.split(cue):
        if p:
            out.append(R(p, tib=True, sz=20, color=RED) if TIB_RE.match(p) else R(p, rstyle='Cue'))
    return out


def para(ln, style, text, cue=None, marks=None, rubric=False, tib=False):
    runs = [R(text)] if tib else latin_runs(text, marks, ln, rubric)
    return P(ln, style, runs + (cue_runs(cue) if cue else []))


# ---------------------------------------------------------------- parsing

def parse_front(lines):
    """Flat `key: value` front matter -> (dict, index of first body line)."""
    meta = {}
    if lines and lines[0].strip() == '---':
        for k in range(1, len(lines)):
            if lines[k].strip() == '---':
                return meta, k + 1
            key, sep, val = lines[k].partition(':')
            if sep:
                meta[key.strip()] = val.strip().strip('"\'')
        err('build', 'front matter is not closed with ---')
    return meta, 0


def title_block(rows):
    """First run of >= 5 blank lines: (cover rows, body-title rows, index after the title)."""
    blanks = 0
    for k, (_, t) in enumerate(rows):
        blanks = blanks + 1 if not t.strip() else 0
        if blanks >= 5:
            s, e = k - 4, k + 1
            while e < len(rows) and not rows[e][1].strip():
                e += 1
            f = e
            while f < len(rows) and rows[f][1].strip():
                f += 1
            return [r for r in rows[:s] if r[1].strip()], rows[e:f], f
    return None


def split_pho(ln, t):
    """Tibetan line with glued Latin text (>= 2 words, e.g. CAPS phonetics) -> [(ln, tibetan), (ln, latin)]."""
    end = max((m.end() for m in TIB_RE.finditer(t)), default=0)
    pho = t[end:].strip()
    if end and sum(1 for w in pho.split() if re.search(r'[^\W\d_]', w)) >= 2 and '>>' not in pho:
        return [(ln, t[:end].rstrip()), (ln, pho)]
    return [(ln, t)]


def groups(rows):
    """rows -> [('sep',) | ('cue', ln, cue) | ('grp', [(ln, text, cue), ...])].

    A Tibetan line starts a group and takes the following non-blank non-Tibetan lines;
    a Latin line with no Tibetan before it is its own group.
    """
    out, cur, blanks = [], None, 0

    def sep(force):
        if (out or force) and not (out and out[-1][0] == 'sep'):
            out.append(('sep',))

    for ln, raw in rows:
        if not raw.strip():
            blanks += 1
            cur = None
            continue
        if blanks >= 2:
            sep(False)
        blanks = 0
        if raw.strip() == '—':
            sep(True)
            cur = None
            continue
        t, cue = split_cue(raw)
        if not t:
            out.append(('cue', ln, cue))
            cur = None
        elif has_tib(t):
            cur = [(ln, t, cue)]
            out.append(('grp', cur))
        elif cur is not None:
            cur.append((ln, t, cue))
        else:
            out.append(('grp', [(ln, t, cue)]))
    return out


def emit(g, last_tib, marks, paras, first=False, nxt=''):
    """Classification rules a-k of the plan; appends paragraphs.

    `first` = first group of a body without title block; `nxt` = first line of the next group.

    Group without Tibetan: CAPS (i) > after the last Tibetan line = Colophon (k) > short (j) > Rubric (k).
    """
    tl = g[0] if has_tib(g[0][1]) else None
    L = g[1:] if tl else list(g)

    def add(row, style, **kw):
        paras.append(para(row[0], style, row[1], row[2], marks, **kw))

    def heading(lvl, tail='Translation'):
        if tl:
            add(tl, f'Heading{lvl}Tib', tib=True)
        add(L[0], f'Heading{lvl}')
        for r in L[1:]:
            add(r, tail, rubric=tail == 'Rubric')

    def heuristic():
        info(f'heading(heuristic) line {L[0][0]}: {L[0][1][:60]}')
        heading(2)

    rub = bool(tl) and tl[1].lstrip('༄༅ \t').startswith(('ཞེས', 'ཅེས'))   # rubric opener, never a heading
    if L and L[0][1].startswith('#'):                                   # a
        lvl = 2 if L[0][1].startswith('##') else 1
        L[0] = (L[0][0], L[0][1].lstrip('#').strip(), L[0][2])
        heading(lvl)
    elif first and tl and L and tl[1].lstrip().startswith('༄'):          # a2 (before b)
        heading(1, 'Rubric')                                            # author lines italic
    elif tl is None:
        t = L[0][1]
        if is_caps(t):                                                  # i
            heading(1)
        elif last_tib and L[0][0] > last_tib:                           # k (colophon, before j)
            add(L[0], 'Colophon')
        elif short(t):                                                  # j
            heuristic()
        else:                                                           # k
            add(L[0], 'Rubric', rubric=True)
    elif not L and (nxt.startswith('#') or H1_TIB_RE.search(tl[1])):   # h: heading-like Tibetan alone
        add(tl, f'Heading{2 if nxt.startswith("##") else 1}Tib', tib=True)
    elif not L:                                                         # h
        add(tl, 'TibVerse', tib=True)
        warn('build', f'line {tl[0]}: Tibetan line without a translation, typed as verse')
    elif len(L) >= 2:
        if is_pho(L[0][1]):                                             # b
            add(tl, 'TibVerse', tib=True)
            add(L[0], 'Phonetics')
            for r in L[1:]:
                add(r, 'Translation')
        else:                                                           # c
            add(tl, 'RubricTib', tib=True)
            for r in L:
                add(r, 'Rubric', rubric=True)
    elif H1_TIB_RE.search(tl[1]) and not rub and len(L[0][1].split()) <= 12:   # d
        heading(1)
    elif (VERSE_END_RE.search(tl[1]) and not rub and 'ནི།' not in tl[1]      # d2
          and not PROSE_END_RE.search(tl[1]) and not MANTRA_IN_RE.search(L[0][1])
          and not is_caps(L[0][1]) and not short(L[0][1]) and not L[0][1].rstrip().endswith(':')):
        info(f'verse(no phonetics) line {tl[0]}: {tl[1][:60]}')
        add(tl, 'TibVerse', tib=True)
        add(L[0], 'Translation')
    elif is_caps(L[0][1]):                                              # e
        add(tl, 'TibVerse', tib=True)
        add(L[0], 'MantraPhonetics')
    elif short(L[0][1]) and not rub:                                    # f
        heuristic()
    else:                                                               # g
        add(tl, 'RubricTib', tib=True)
        add(L[0], 'Rubric', rubric=True)


# ---------------------------------------------------------------- cover image

def image_size(data):
    """(width, height) in px from PNG IHDR or JPEG SOF."""
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return struct.unpack('>II', data[16:24])
    if data[:2] == b'\xff\xd8':
        i = 2
        while i + 4 <= len(data) and data[i] == 0xFF:
            mk, ln = data[i + 1], struct.unpack('>H', data[i + 2:i + 4])[0]
            if 0xC0 <= mk <= 0xCF and mk not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack('>HH', data[i + 5:i + 9])
                return w, h
            i += 2 + ln             # ponytail: no 0xFF fill-byte padding between segments
    raise ValueError('not a PNG/JPEG file')


def drawing_xml(cx, cy, name):
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="Cover image"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{escape(name)}"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId6"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


# ---------------------------------------------------------------- styles

# (id, name, basedOn, kind, pPr kwargs, rPr kwargs); sizes in half-points, spacing in twips
STYLES = [
    ('Normal', 'Normal', None, 'p', dict(tabs=True), {}),
    ('TibVerse', 'TibVerse', 'Normal', 'p', dict(keep=True), dict(tib=True, sz=38, color=RED)),
    ('Phonetics', 'Phonetics', 'Normal', 'p', dict(keep=True), dict(b=True, sz=24)),
    ('MantraPhonetics', 'MantraPhonetics', 'Phonetics', 'p', dict(keep=False, after=140), {}),
    ('Translation', 'Translation', 'Normal', 'p', dict(after=140), dict(sz=24)),
    ('Separator', 'Separator', 'Normal', 'p', dict(before=140), dict(color=RED)),
    ('RubricTib', 'RubricTib', 'Normal', 'p', dict(keep=True), dict(tib=True, sz=28)),
    ('Rubric', 'Rubric', 'Normal', 'p', dict(after=140), dict(i=True, sz=22)),
    ('Heading1Tib', 'Heading1Tib', 'Normal', 'p', dict(keep=True), dict(tib=True, sz=36)),
    ('Heading1', 'heading 1', 'Normal', 'p', dict(keep=True, after=320, outline=0),
     dict(b=True, caps=True, color=RED, sz=28)),
    ('Heading2Tib', 'Heading2Tib', 'Normal', 'p', dict(keep=True), dict(tib=True, sz=28)),
    ('Heading2', 'heading 2', 'Normal', 'p', dict(after=160, outline=1),
     dict(b=True, caps=True, color=RED, sz=24)),
    ('Cue', 'Cue', None, 'c', {}, dict(b=True, sz=20, color=RED)),
    ('Drum', 'Drum', None, 'c', {}, dict(color=BLUE)),
    ('EndnoteReference', 'endnote reference', None, 'c', {}, dict(sup=True)),
    ('EndnoteText', 'endnote text', 'Normal', 'p', dict(after=140), dict(sz=22)),
    ('Header', 'header', 'Normal', 'p', dict(jc='right'), dict(b=True, sz=20)),
    ('Footer', 'footer', 'Normal', 'p', {}, dict(b=True, sz=20, color=RED)),
    ('CoverTib', 'CoverTib', 'Normal', 'p', dict(jc='center'), dict(tib=True, sz=48)),
    ('CoverTitle', 'CoverTitle', 'Normal', 'p', dict(jc='center'),
     dict(b=True, caps=True, color=RED, sz=44)),
    ('CoverSub', 'CoverSub', 'Normal', 'p', dict(jc='center'), dict(sz=28)),
    ('CoverImage', 'CoverImage', 'Normal', 'p', dict(jc='center'), {}),
    ('CoverBottomTib', 'CoverBottomTib', 'Normal', 'p', dict(jc='center'), dict(tib=True, sz=38)),
    ('CoverBottom', 'CoverBottom', 'Normal', 'p', dict(jc='center'), dict(b=True, sz=24)),
    ('Colophon', 'Colophon', 'Normal', 'p', dict(after=140), dict(sz=22)),
    ('NotesLabel', 'NotesLabel', 'Normal', 'p', dict(before=280), dict(b=True, sz=24)),
]


def styles_xml():
    f = 'Cambria'
    out = (XML + f'<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>'
           f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
           '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="cs-CZ" w:bidi="bo-CN"/>'
           '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
           '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>')
    for sid, name, base, kind, pk, rk in STYLES:
        typ = 'paragraph' if kind == 'p' else 'character'
        dflt = ' w:default="1"' if sid == 'Normal' else ''
        out += (f'<w:style w:type="{typ}"{dflt} w:styleId="{sid}"><w:name w:val="{name}"/>'
                + (f'<w:basedOn w:val="{base}"/>' if base else '')
                + '<w:qFormat/>' + ppr(**pk) + rpr(**rk) + '</w:style>')
    return out + '</w:styles>'


# ---------------------------------------------------------------- package

def endnotes_xml(notes):
    sep = ('<w:endnote w:type="{t}" w:id="{i}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
           'w:lineRule="auto"/></w:pPr><w:r><w:{t}/></w:r></w:p></w:endnote>')
    out = XML + f'<w:endnotes {NS}>' + sep.format(t='separator', i=-1) \
        + sep.format(t='continuationSeparator', i=0)
    for n, text in enumerate(notes, 1):
        out += (f'<w:endnote w:id="{n}"><w:p>{ppr("EndnoteText")}'
                '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>'
                f'<w:r><w:t xml:space="preserve"> </w:t></w:r>{runs_xml(latin_runs(text))}</w:p></w:endnote>')
    return out + '</w:endnotes>'


def write_docx(out, paras, notes, meta, img, titlepg):
    """img = (bytes, ext, cx, cy) or None."""
    has_h = bool(meta.get('header1') or meta.get('header2'))
    ct = ('application/vnd.openxmlformats-officedocument.wordprocessingml.', '+xml')
    over = [('/word/document.xml', 'document.main'), ('/word/styles.xml', 'styles'),
            ('/word/settings.xml', 'settings'), ('/word/endnotes.xml', 'endnotes'),
            ('/word/footer1.xml', 'footer')]
    if has_h:
        over.append(('/word/header1.xml', 'header'))
    types = (XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/>'
             '<Default Extension="png" ContentType="image/png"/>'
             + ''.join(f'<Override PartName="{n}" ContentType="{ct[0]}{t}{ct[1]}"/>' for n, t in over)
             + '</Types>')
    root_rels = (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>')
    rels = [('rId1', 'styles', 'styles.xml'), ('rId2', 'settings', 'settings.xml'),
            ('rId3', 'endnotes', 'endnotes.xml'), ('rId5', 'footer', 'footer1.xml')]
    if has_h:
        rels.append(('rId4', 'header', 'header1.xml'))
    if img:
        rels.append(('rId6', 'image', f'media/cover{img[1]}'))
    doc_rels = (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + ''.join(f'<Relationship Id="{i}" Type="{REL}/{t}" Target="{g}"/>' for i, t, g in rels)
                + '</Relationships>')
    sect = ('<w:sectPr>'
            + ('<w:headerReference w:type="default" r:id="rId4"/>' if has_h else '')
            + '<w:footerReference w:type="default" r:id="rId5"/>'
            '<w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="720" '
            'w:footer="709" w:gutter="0"/><w:cols w:space="708"/>'
            + ('<w:titlePg/>' if titlepg else '') + '<w:docGrid w:linePitch="360"/></w:sectPr>')
    document = (XML + f'<w:document {NS} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>'
                + ''.join(p_xml(p) for p in paras) + sect + '</w:body></w:document>')
    settings = (XML + f'<w:settings {NS}><w:defaultTabStop w:val="708"/>'
                '<w:characterSpacingControl w:val="doNotCompress"/>'
                '<w:endnotePr><w:numFmt w:val="lowerRoman"/><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    footer = (XML + f'<w:ftr {NS}><w:p>{ppr("Footer")}{runs_xml(latin_runs(meta.get("footer", "")))}'
              '<w:r><w:tab/></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t xml:space="preserve">1</w:t></w:r></w:fldSimple></w:p></w:ftr>')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles_xml())
        z.writestr('word/settings.xml', settings)
        z.writestr('word/endnotes.xml', endnotes_xml(notes))
        z.writestr('word/footer1.xml', footer)
        if has_h:
            hp = ''.join(f'<w:p>{ppr("Header")}{runs_xml(latin_runs(meta[k]))}</w:p>'
                         for k in ('header1', 'header2') if meta.get(k))
            z.writestr('word/header1.xml', XML + f'<w:hdr {NS}>{hp}</w:hdr>')
        if img:
            z.writestr(f'word/media/cover{img[1]}', img[0])


# ---------------------------------------------------------------- build

def find_title(body, tb):
    """Document title for the default header2: body title block, else cover, else yig mgo title group, else first `#` line."""
    def clean(t):
        t = re.sub(r'(?<=[,.;:!?…)\]"“”»’])[ivxlc]{1,4}(?=\s|$)', '', demark(split_cue(t)[0]))
        return t.lstrip('#').strip()
    cands = [t for rows in (tb[1], tb[0]) for _, t in rows if not has_tib(t)] if tb else []
    nb = [t for _, t in body if t.strip()]
    if not tb and nb and nb[0].lstrip().startswith('༄'):   # yig mgo title group (rule a2)
        cands += [t for t in nb[1:2] if not has_tib(t)]
    cands += [t for _, t in body if re.match(r'#(?!#)', t)]
    return next((c for c in map(clean, cands) if c), '')


def resolve_headers(meta, cli, title):
    """CLI (None = not given; '' suppresses) > front matter > default (header2 = title)."""
    h = {k: meta.get(k, '') for k in ('header1', 'header2', 'footer')}
    if 'header2' not in meta:
        h['header2'] = title
    h.update({k: v for k, v in (cli or {}).items() if v is not None})
    return h


def build(src, out, dry, cli=None):
    lines = Path(src).read_text(encoding='utf-8').split('\n')
    meta, fm = parse_front(lines)
    global TIB_FONT
    TIB_FONT = meta.get('tib_font', 'Jomolhari')
    label = meta.get('notes_label', 'Poznámky:')
    rows = [(i + 1, lines[i]) for i in range(fm, len(lines))]
    rows = [r for ln, t in rows for r in split_pho(ln, t)]

    ni = next((k for k in range(len(rows) - 1, -1, -1)
               if rows[k][1].strip() == label or rows[k][1].startswith(label)), None)
    notes, raw_notes, body = [], [], rows
    if ni is not None:
        first_note = rows[ni][1][len(label):].strip() if rows[ni][1].startswith(label) else ''
        notes = ([first_note] if first_note else []) + [t.strip() for _, t in rows[ni + 1:] if t.strip()]
        body = rows[:ni]
        if notes and re.match(r'i[\s.)]', notes[0]):    # inline numerals in one line: split in strict sequence
            t, parts, k = notes[0], [], 2
            while True:
                m = re.compile(r'\s(?=%s[\s.)])' % roman(k)).search(t)
                if not m:
                    break
                parts.append(t[:m.start()].strip())
                t, k = t[m.end():], k + 1
            notes[:1] = parts + [t.strip()]
        raw_notes = notes
        if notes and re.match(r'i[\s.)]', notes[0]):     # strip leading numerals that match their index
            def strip_num(k, t):
                m = re.match(r'([ivxlc]+)[\s.)]+', t)
                return t[m.end():] if m and m.group(1) == roman(k) else t
            notes = [strip_num(k, t) for k, t in enumerate(notes, 1)]

    paras, marks, img = [], [], None
    tb = title_block(body)
    if tb is None:
        warn('build', 'no title block (a run of >= 5 blank lines): no cover, no body title')
        rest = body
    else:
        cover, title, k = tb
        rest = body[k:]
        for ln, t in cover:
            tib = has_tib(t)
            paras.append(para(ln, 'CoverTib' if tib else 'CoverTitle' if is_caps(t) else 'CoverSub',
                              t.strip(), marks=marks, tib=tib))
        if meta.get('cover_image'):
            path = Path(src).resolve().parent / meta['cover_image']
            try:
                data = path.read_bytes()
                w, h = image_size(data)
                s = min(COVER_BOX[0] / w, COVER_BOX[1] / h)
                cx, cy = int(w * s), int(h * s)
                img = (data, path.suffix.lower(), cx, cy)
                paras.append(P(0, 'CoverImage', [R('[cover image]', raw=drawing_xml(cx, cy, path.name))]))
            except (OSError, ValueError) as e:
                err('build', f'cover_image {path}: {e}')
        gap = True
        for style, key in (('CoverBottomTib', 'cover_tib'), ('CoverBottom', 'cover_publisher'),
                           ('CoverBottom', 'cover_year')):
            if meta.get(key):
                runs = [R(meta[key])] if key == 'cover_tib' else latin_runs(meta[key])
                paras.append(P(0, style, runs, gap=gap))
                gap = False
        first = True
        for ln, t in title:
            tib = has_tib(t)
            style = 'Heading1Tib' if tib else 'Heading1' if first else 'Translation'
            first = first and tib
            p = para(ln, style, t.strip(), marks=marks, tib=tib)
            paras.append(p._replace(pbb=True) if ln == title[0][0] else p)

    last_tib = max((ln for ln, t in rest if has_tib(t)), default=0)
    gs = groups(rest)
    for k, g in enumerate(gs):
        if g[0] == 'sep':
            paras.append(P(0, 'Separator', [R('—')]))
        elif g[0] == 'cue':
            paras.append(P(g[1], None, cue_runs(g[2])))
        else:
            nxt = gs[k + 1][1][0][1] if k + 1 < len(gs) and gs[k + 1][0] == 'grp' else ''
            emit(g[1], last_tib, marks, paras, first=(k == 0 and tb is None), nxt=nxt)
    if ni is not None:
        paras.append(P(rows[ni][0], 'NotesLabel', plain_runs(label)))

    if len(marks) != len(notes):
        hint = ''
        if any(c in runs_text(p.runs) for p in paras for c in '¹²³⁴⁵⁶⁷⁸⁹⁰'):
            hint = '\n  (superscript digits found: old-format text, not supported)'
        warn('build', f'{len(marks)} endnote markers but {len(notes)} notes; notes typeset as plain text\n'
                      f'  markers: {", ".join(f"{m} (line {ln})" for ln, m in marks) or "-"}\n'
                      f'  notes:   {"; ".join(n[:30] for n in notes) or "-"}{hint}')
        # ponytail: markers become plain text as written (`,ii` / `^v`), notes keep their numerals
        paras = [p._replace(runs=[R(r.text) if r.raw and 'endnoteReference' in r.raw else r
                                  for r in p.runs]) for p in paras]
        paras += [P(0, 'Colophon', latin_runs(t)) for t in raw_notes]
        notes = []
    title = find_title(body, tb) or next((runs_text(p.runs) for p in paras if p.style == 'Heading1'), '')
    meta.update(resolve_headers(meta, cli, title))
    for k in ('header1', 'header2', 'footer'):
        info(f'{k}: {meta[k]}')
    if dry:
        for p in paras:
            print(f'{p.lineno}\t{p.style or "Normal"}\t{runs_text(p.runs)[:60]}')
    finish('build')
    if not dry:
        write_docx(out, paras, notes, meta, img, tb is not None)
        print(f'{out}: {len(paras)} paragraphs, {len(notes)} endnotes')
    return paras


# ---------------------------------------------------------------- selftest

EXPECTED = (
    ['CoverTib', 'CoverTitle', 'CoverSub', 'CoverImage', 'CoverBottomTib', 'CoverBottom', 'CoverBottom',
     'Heading1Tib', 'Heading1', 'Translation',                  # body title
     'Heading1Tib', 'Heading1',                                 # H1 via ནི།
     'RubricTib', 'Rubric', 'TibVerse', 'MantraPhonetics',      # rubric, mantra
     'Heading2',                                                # short Latin line
     'TibVerse', 'Phonetics', 'Translation', 'Normal',          # verse 1, cue-only
     'TibVerse', 'Phonetics', 'Translation',                    # verse 2
     'Separator', 'Heading1', 'Separator',                      # 2 blanks, `#`, `—`
     'RubricTib', 'Rubric', 'Colophon', 'NotesLabel'])


def marks_of(s, nxt):
    marks = [None] * (nxt - 1)
    latin_runs(s, marks)
    return len(marks) == nxt


def cmd_selftest(_args):
    fixture = Path(__file__).resolve().parent / 'fixture' / 'text.md'
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / 'out.docx'
        build(fixture, out, False)
        z = zipfile.ZipFile(out)
        doc = z.read('word/document.xml').decode('utf-8')
        endn = z.read('word/endnotes.xml').decode('utf-8')
        hdr = z.read('word/header1.xml').decode('utf-8')
    assert 'Překlad pro vnitřní použití' in hdr, 'fixture header2 lost'
    with tempfile.TemporaryDirectory() as td:
        md, out = Path(td) / 't.md', Path(td) / 'o.docx'
        md.write_text('# TITLE\nཀ\nKA\nA verse.\n', encoding='utf-8')
        build(md, out, False)
        assert 'TITLE' in zipfile.ZipFile(out).read('word/header1.xml').decode('utf-8'), 'default header2'
    assert resolve_headers({'header2': 'FM'}, {'header2': 'CLI', 'footer': None}, 'T')['header2'] == 'CLI'
    assert resolve_headers({'header2': 'FM'}, {'header2': None}, 'T')['header2'] == 'FM'
    assert resolve_headers({}, {'header2': ''}, 'T')['header2'] == ''
    assert resolve_headers({}, {}, 'T')['header2'] == 'T'
    assert find_title([(1, '༄༅། །ཚོགས།'), (2, 'DENNÍ OBĚTINA'), (3, 'složil X')], None) == 'DENNÍ OBĚTINA'
    seq = [(re.search(r'<w:pStyle w:val="(\w+)"/>', p) or [None, 'Normal'])[1]
           for p in re.findall(r'<w:p>(.*?)</w:p>', doc)]
    assert seq == EXPECTED, f'pStyle sequence differs:\n{seq}\n{EXPECTED}'
    assert len(re.findall(r'<w:endnote w:id="\d+">', endn)) == 3, 'expected 3 endnotes'
    assert re.search(r'<w:sectPr>.*lowerRoman.*</w:sectPr>', doc, re.S), 'lowerRoman missing in sectPr'
    assert doc.count('<w:endnoteReference') == 3, 'expected 3 endnote references'
    text = re.findall(r'<w:t xml:space="preserve">(.*?)</w:t>', doc, re.S)
    for bad in ('{', '}', '>>', '&gt;&gt;', '*'):
        assert not any(bad in t for t in text), f'leftover {bad!r} in w:t'
    assert [roman(n) for n in range(1, 16)] == \
        'i ii iii iv v vi vii viii ix x xi xii xiii xiv xv'.split(), 'roman'
    assert marks_of('moudrosti,ii', 2) and marks_of('STEZKYi', 1) and marks_of('slovo^v', 5)
    assert not marks_of('siddhi', 1) and not marks_of('v Praze', 5) and not marks_of('(i)', 1)
    assert short('Dar Dharmy') and not short('Recitujte třikrát') and not short('Takto ji nabídněte.')
    assert not short('*složil Düdžom Rinpočhe*') and short('*Dar* Dharmy')
    paras = []
    emit([(5, '*složil Düdžom Rinpočhe*', None)], 9, [], paras)
    emit([(10, 'Zdroj: [dudjom.tsadra.org](http://dudjom.tsadra.org) (DJYD-1)', None)], 9, [], paras)
    emit([(11, 'DAR', None)], 9, [], paras)
    assert [p.style for p in paras] == ['Rubric', 'Colophon', 'Heading1'], [p.style for p in paras]
    assert runs_text(paras[1].runs) == 'Zdroj: dudjom.tsadra.org (DJYD-1)'
    assert runs_text(latin_runs('[Web](https://x.org/a) a [x.org](x.org/a/b)')) == 'Web (https://x.org/a) a x.org'
    def styles(*g, **kw):
        ps = []
        emit(list(g), 99, [], ps, **kw)
        return [p.style for p in ps]
    tb, la = (1, 'ཀ།', None), (2, 'Název díla', None)
    assert styles(tb, nxt='## Titul') == ['Heading2Tib'] and styles(tb, nxt='# T') == ['Heading1Tib']   # 1
    assert styles((1, 'ཚེ་འགུགས་ནི།', None)) == ['Heading1Tib']
    assert styles((1, '༄༅། །ཀ', None), la, (3, 'Autor', None), first=True) == \
        ['Heading1Tib', 'Heading1', 'Rubric']                                                         # 2
    assert styles((1, '༄༅། །ཀ', None), (2, 'PŘIVOLÁNÍ VĚDOMÍ', None), first=True)[1] == 'Heading1'
    assert split_pho(7, 'ཀ༔ ČHI NANG SANG') == [(7, 'ཀ༔'), (7, 'ČHI NANG SANG')]         # 3
    assert split_pho(7, 'ཀ༔ OM') == [(7, 'ཀ༔ OM')]
    assert split_pho(7, 'ཀ། ། Chcete-li provádět praxi') == [(7, 'ཀ། །'), (7, 'Chcete-li provádět praxi')]
    assert split_pho(7, 'ཀ། Text >> cue words') == [(7, 'ཀ། Text >> cue words')]
    assert not short('složil Düdžom Rinpočhe')                                                        # 4
    assert styles((1, 'ཞེས་ཚོགས་པའི་མཐར།', None), (2, 'Na konci zásluh', None)) == ['RubricTib', 'Rubric']  # 5
    assert styles((1, 'ཀ། །', None), (2, 'přijměte prosím obětiny.', None)) == ['TibVerse', 'Translation']
    assert styles((1, 'ཞེས་ཀ། །', None), (2, 'Toto složil Džigdräl.', None)) == ['RubricTib', 'Rubric']
    assert styles((1, 'ཚིག་བདུན་གསོལ་འདེབས་ནི།', None), (2, 'Sedmiřádková modlitba', None))[0] == 'Heading1Tib'
    assert styles((1, 'ཀ་ནི།', None), (2, ' '.join(['slovo'] * 13), None))[0] == 'RubricTib'       # 6
    with tempfile.TemporaryDirectory() as td:                                                         # inline notes
        md, out = Path(td) / 't.md', Path(td) / 'o.docx'
        md.write_text('ཀ\nKA\nA verse,i\nB,ii\nC,iii\n\nPoznámky: i Prvá. ii Druhá, viz ii. iii Třetí\n', encoding='utf-8')
        build(md, out, False)
        assert zipfile.ZipFile(out).read('word/endnotes.xml').decode('utf-8').count('<w:endnote w:id="') == 3, 'inline notes split'
    with tempfile.TemporaryDirectory() as td:                                                         # 7
        md = Path(td) / 't.md'
        md.write_text('ཀ\nKA\nA verse,i\n\nPoznámky: První\n', encoding='utf-8')
        build(md, None, True)
    with tempfile.TemporaryDirectory() as td:                                                         # 8
        md, out = Path(td) / 't.md', Path(td) / 'o.docx'
        md.write_text('ཀ\nKA\nA verse,i\n\nPoznámky:\ni První\nii Druhá\n', encoding='utf-8')
        with contextlib.redirect_stderr(io.StringIO()) as e:
            build(md, out, False)
        assert 'notes typeset as plain text' in e.getvalue(), 'mismatch must warn'
        z = zipfile.ZipFile(out)
        doc = z.read('word/document.xml').decode('utf-8')
        assert '<w:endnoteReference' not in doc and 'A verse,i' in re.sub('<[^>]+>', '', doc), 'markers must stay literal'
        assert '<w:endnote w:id="1">' not in z.read('word/endnotes.xml').decode('utf-8')
        seq = re.findall(r'<w:pStyle w:val="(\w+)"/>', doc)
        assert seq[-3:] == ['NotesLabel', 'Colophon', 'Colophon'], seq
        assert 'ii Druhá' in doc and 'i První' in doc, 'note numerals preserved'
    assert PROSE_END_RE.search('དགེའོ།། །།')          # དགེའོ།། །།
    assert PROSE_END_RE.search('བསྔོ་བྱའོ། །')          # བྱའོ། །
    assert not PROSE_END_RE.search('བཞེས་སུ་གསོལ། །')  # གསོལ། །
    assert MANTRA_IN_RE.search('s pomocí DZA HUNG BAM HO si představte')
    assert not MANTRA_IN_RE.search('přijměte prosím tyto čisté obětiny.')
    print('selftest OK')
    return 0


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('build')
    p.add_argument('text')
    p.add_argument('-o', '--out')
    p.add_argument('--dry-run', action='store_true')
    for k in ('header1', 'header2', 'footer'):
        p.add_argument(f'--{k}')
    p.set_defaults(fn=lambda a: build(a.text, a.out, a.dry_run,
                                      {k: getattr(a, k) for k in ('header1', 'header2', 'footer')}) and 0)
    p = sub.add_parser('selftest')
    p.set_defaults(fn=cmd_selftest)
    args = ap.parse_args()
    if args.cmd == 'build' and not args.out and not args.dry_run:
        ap.error('build: -o/--out is required (or --dry-run)')
    sys.exit(args.fn(args) or 0)


if __name__ == '__main__':
    main()
