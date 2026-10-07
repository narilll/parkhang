# Parkhang

A Claude Code skill that typesets an interlinear `text.md` — Tibetan line, UPPERCASE
phonetics, translation — into a Word `.docx` in a fixed print style: A4, Cambria for the
Latin text, Jomolhari for the Tibetan, red Tibetan verses, red caps headings, italic
rubrics, right-aligned cue labels, real Word endnotes with lowercase roman numbering, an
optional cover page. The input is the output format of the sibling
lotsawa translation skill, but any text in that layout works.
Parkhang (པར་ཁང་) is Tibetan for a printing house.

The script is Python 3.9 standard library only and writes the OOXML parts directly. No
pip or npm packages.

## Contents

| File | Purpose |
|---|---|
| `SKILL.md` | Skill definition — input contract, classification rules, steps, style table |
| `scripts/parkhang.py` | `build` (md → docx) and `selftest` |
| `scripts/fixture/text.md`, `cover.jpg` | Test document covering every element |

## Install

Copy or clone the folder into your skills directory (`~/.claude/skills/parkhang` or a
project's `skills/parkhang`). Claude Code picks it up at the start of the next session.

Optional, for render checks only: LibreOffice and the Jomolhari font
(`brew install --cask libreoffice font-jomolhari`). Recipients open the `.docx` in Word
with Cambria and Jomolhari installed.

## Usage

```
python3 scripts/parkhang.py build text.md --dry-run          # shows the style per line
python3 scripts/parkhang.py build text.md -o out.docx \
    --header1 "DÜDŽOM TERSAR / KHANDRO THUGTHIG" --footer "ORGYEN KHANDRO LING"
python3 scripts/parkhang.py selftest
```

Or in Claude Code: `/parkhang <path to text.md>`.

## Input

Plain text, one line per element, no markup required:

```
༄༅། །ཚོགས་བསྡུས་ནི།
STRUČNÁ OBĚTINA TSOKU
složil Mipham Rinpočhe

<about 18 blank lines — the cover page>

༄༅། །ཚོགས་བསྡུས་ནི།
Stručná obětina tsoku
složil Mipham Rinpočhe

ཚོགས་རྫས་རྣམས། ཨོཾ་ཨཱཿཧཱུྃ་ཧོས། བྱིན་གྱིས་རླབས།
Substance tsoku požehnejte pomocí OM AH HUNG HO.

ཕྱི་ནང་གསང་བའི་བདེ་ཆེན་ཚོགས་མཆོད་འབུལ། །
ČHI NANG SANG WA'I DE ČHEN TSHOG ČHÖ BÜL
Přináším obětinu tsoku velké blaženosti a moudrosti,ii vnější i vnitřní.
```

| Lines | Element |
|---|---|
| Tibetan + UPPERCASE + translation | verse |
| Tibetan + UPPERCASE | mantra |
| Tibetan + mixed-case line | rubric (instruction) — or heading, see below |
| Latin line alone | heading if short, otherwise rubric; after the last Tibetan line: colophon |
| `Poznámky:` + lines | endnotes; markers in the text are roman numerals attached to a word (`moudrosti,ii`, `SLOVOi`, `slovo^v`) |

Heading vs. rubric is decided heuristically: a Tibetan line ending in the topic marker
`ནི།` is a heading; a short Latin line with no final punctuation and no imperative is a
heading. `--dry-run` shows every decision. A `#`/`##` prefix forces a heading, a trailing
`.` forces a rubric.

### Extensions

| Syntax | Effect |
|---|---|
| YAML-like front matter `---` … `---` | `header1`, `header2` (default: document title), `footer` (text left of the page number), `cover_image`, `cover_tib`, `cover_publisher`, `cover_year`, `tib_font` (default `Jomolhari`), `notes_label` (default `Poznámky:`) |
| `--header1 / --header2 / --footer` | same, from the command line; `""` suppresses a line |
| ` >> MELODIE` at line end | cue label flush right, red bold 10 pt; Tibetan ornaments (࿂ ࿃ ࿄) allowed |
| `{DÜN}` in phonetics | blue syllable (drum beat) |
| `*ja la la*` | italic |
| `—` alone, or two blank lines | separator paragraph |
| `[text](url)` | flattened to `text` or `text (url)` |

## Style

Measured from the reference document (Word export, A4, margins 2.5 / 2.5 / 2.5 / 2.0 cm,
one right tab at 16 cm):

| Element | Font | Size | Colour |
|---|---|---|---|
| Tibetan verse | Jomolhari | 19 | `C00000` |
| Phonetics | Cambria Bold caps | 12 | black |
| Translation | Cambria | 12 | black, 7 pt after |
| Tibetan rubric / rubric | Jomolhari 14 / Cambria Italic 11 | | black; mantra words bold upright |
| Heading 1 / 2 | Cambria Bold caps | 14 / 12 | `C00000`, Tibetan line above in black 18 / 14 |
| Header / footer | Cambria Bold | 10 | black / `C00000` |
| Cue label | Cambria Bold | 10 | `C00000` |
| Drum syllable | inherits | | `007BB8` |
| Endnotes | Cambria | 11 | lowercase roman |
| Cover title | Jomolhari 24 / Cambria Bold caps 22 | | black / `C00000`, centred |

Everything is a named Word style (`TibVerse`, `Phonetics`, `Translation`, `Rubric`,
`Heading1`, `Cue`, …), so the whole document can be restyled in Word's Styles pane.

## Checking the output

```
soffice --headless --convert-to pdf --outdir /tmp/out out.docx
pdftoppm -png -r 70 /tmp/out/out.pdf /tmp/out/p
```

LibreOffice on macOS may see only its bundled fonts in headless mode; if the Tibetan
renders blank, copy the font into the bundle:
`cp ~/Library/Fonts/Jomolhari-Regular.ttf /Applications/LibreOffice.app/Contents/Resources/fonts/truetype/`.
Cambria is substituted by the metric-compatible Caladea.

## Limits

- Old interlinear texts with lowercase phonetics, an extra IAST line or `¹` superscripts
  are rejected (marker/note mismatch), not mis-typeset.
- The cover bottom block sits 36 pt below the image instead of being pinned to the page
  foot; pin it by hand in Word if wanted.
- `header1` (cycle name) cannot be derived from the text; pass it explicitly.
