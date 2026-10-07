---
name: parkhang
description: Use when the user wants a Word document (.docx) typeset from an interlinear text.md (lotsawa output: Tibetan / UPPERCASE phonetics / translation) in the house print style — cover page, red Tibetan, endnotes. Input is a text.md; output is a .docx only.
---

# Parkhang — interlinear text.md → Word .docx

`scripts/parkhang.py` does all the mechanics: it classifies every segment of a lotsawa `text.md`, then writes the `.docx`
as raw OOXML (Python stdlib only, no pip/npm). **Never write OOXML by hand and never post-process the `.docx` yourself** —
fix the `text.md` (or the script) and rebuild. The deliverable is the `.docx` only; a PDF is rendered solely so you can
check the layout, and is not handed over.

## Input

A lotsawa `text.md` (contract in the lotsawa skill, "Output format"):

- **Title block** — cover lines, a run of >= 5 blank lines, then the body title (up to the next blank line).
- **Verse** — 3 lines: Tibetan, UPPERCASE phonetics, translation.
- **Rubric** — 2 lines: Tibetan, translation (no phonetics). **Mantra** — 2 lines: Tibetan, UPPERCASE phonetics.
- **Heading** — a translated line on its own, plus the Tibetan line when the source has one.
- **Colophon** — Tibetan + translation (typeset as a rubric: Jomolhari 14 black + italic 11), then the
  translator-credit paragraph (`Colophon` style). Without front matter the document has no header, no cover image and no
  cover bottom block; the cover is the title block alone and the footer carries only the page number.
- **Notes** — the last line equal to (or starting with) the notes label (default `Poznámky:`); text after the label on that
  line is the first note. Then one note per line, each opening with its roman numeral (`i`, `ii`, ...).

Optional extensions (all invisible to a plain lotsawa text, add only what you need):

| Extension | Where | Effect |
|---|---|---|
| Front matter `---` … `---` | top of file, flat `key: value` | `header1`, `header2` (header lines, right-aligned; `header2` defaults to the document title: body title, else cover title, else first `#` line; `header1` has no default, pass the cycle name via front matter or `--header1`; no header text at all = no header), `footer` (text before the page number), `cover_image` (path relative to the md), `cover_tib`, `cover_publisher`, `cover_year` (bottom block of the cover), `tib_font` (default `Jomolhari`), `notes_label` (default `Poznámky:`) |
| ` >> CUE` | end of any line | Cue label flush right, red bold (e.g. `>> MELODIE`) |
| `>> CUE` alone | own line | Cue-only paragraph (style `Normal`), e.g. `>> ࿂࿄ / ZILŇEN ࿃🥁`; Tibetan ornaments in the cue use the Tibetan font, 10 pt red |
| `{SYL}` | phonetics | Blue drum syllable (`{DÜN}`) |
| `*x*` | Latin text | Italic (`*ja la la*`) |
| `#` / `##` prefix | Latin line | Forces heading level 1 / 2 (the `#` is stripped) |
| `—` alone on a line, or 2+ blank lines | between segments | Separator paragraph (red `—`); one blank line = nothing |
| Endnote markers `slovo,ii`, `SLOVOi`, `slovo^v` | Latin text, also cover/title | Real Word endnote reference |

Example (from `scripts/fixture/text.md`):

```
---
header1: Dlouhá praxe Ješe Tsogjäl
cover_image: cover.jpg
---
ཕྱི་ནང་གསང་བའི་བདེ་ཆེན་ཚོགས་མཆོད་འབུལ། །
ČHI NANG SANG WA’I {DÜN} DE ČHEN TSHOG ČHÖ BÜL >> MELODIE
Přináším obětinu tsoku velké blaženosti a moudrosti,ii vnější i vnitřní.
>> ࿂࿄ / ZILŇEN ࿃🥁
```

Marker rules: a roman numeral counts as a marker only (1) after punctuation `,.;:!?…)]"“”»’`, an uppercase letter, or an
explicit `^`, AND (2) when it is exactly the next numeral in sequence (`i`, then `ii`, ...). Anything else (`siddhi`,
`v Praze`, `(i)`) stays literal text. The number of accepted markers must equal the number of notes.

- Run lotsawa `check` **before** adding extensions: ornaments on Latin lines trip its Tibetan-character check and front
  matter confuses it.
- **Old-format texts are not supported** (lowercase phonetics, IAST line, `¹ ²` superscripts): the build fails with the
  marker/note mismatch error (it adds a hint when superscript digits are present). Do not work around it.
- No title block (no run of >= 5 blank lines) = warning, no cover, no body title page break, footer on every page.

## Classification

Line types: `TIB` = contains U+0F00–0FFF; `PHO` = single-case letters and none of `, ! ? ; :` (same test as lotsawa
`looks_like_pho`); `CAPS` = all letters uppercase. `{ } *` and endnote markers are ignored in these checks. A trailing
` >> CUE` is cut off first. A `TIB` line starts a group and takes the following non-blank non-`TIB` lines (`L`); a Latin
line with no Tibetan before it is its own group. First matching rule wins:

| # | Condition | Result |
|---|---|---|
| a | `L[0]` starts `#` / `##` | Heading1 / Heading2 (+ Tibetan line as Heading1Tib / Heading2Tib) |
| a2 | first group of a body without title block, starts with `༄`, has Latin lines | Heading1Tib + Heading1 (first Latin line) + Translation for the rest (wins over b) |
| b | TIB, `len(L)>=2`, `PHO(L[0])` | Verse: TibVerse, Phonetics, Translation for the rest |
| c | TIB, `len(L)>=2`, not PHO | RubricTib + Rubric per line |
| d | TIB, 1 Latin line of <= 12 words, Tibetan ends `ནི` + `།`/`༔` | Heading1 (+ Heading1Tib) |
| e | TIB, 1 Latin line, `CAPS` | Mantra: TibVerse + MantraPhonetics |
| f | TIB, 1 Latin line, `short()` | Heading2 (+ Heading2Tib) — heuristic, logged |
| | (d, e-f) Tibetan starting `ཞེས` / `ཅེས` (after `༄༅`) | never a heading: falls to g |
| g | TIB, 1 Latin line | RubricTib + Rubric |
| h | TIB, no Latin line, next group starts `#` / `##` or the line ends `ནི།` | Heading1Tib / Heading2Tib (matching the `#` level; `ནི།` = level 1), no warning |
| h | TIB, no Latin line, otherwise | TibVerse + warning |
| i | no TIB, `CAPS` | Heading1 |
| j | no TIB, after the last Tibetan line of the text | Colophon |
| k | no TIB, `short()` | Heading2 — heuristic, logged |
| l | no TIB | Rubric |

`short()` is also false for a line wrapped entirely in `*…*` (an italic author line is a rubric, not a heading) and for a
line whose first letter is lowercase (a continuation such as `složil …`).
Preprocessing: a `TIB` line followed on the same line by >= 2 CAPS Latin words (glued phonetics) is split into the
Tibetan line and a phonetics line.
Markdown links `[text](url)` become `text` when one contains the other, else `text (url)`.

`short(s)` = at most 8 words, first letter not lowercase, last character not in `.:!,;?`, and no Czech 2nd-person-plural imperative (`\w+te`, `ete`,
`ěte`, `jte`). It is Czech-only; for another target language expect misfires.

In Rubric paragraphs, tokens of 2+ letters that are all uppercase (`OM AH HUNG HO`) become bold non-italic.

**Fixing a misfire** (edit the `text.md`, rebuild): a trailing `.` makes `short()` fail, so the line becomes a rubric; a
`#` or `##` prefix forces a heading. `--dry-run` shows the decision per line and prints `INFO heading(heuristic) line N: ...`
for every heading decided by `short()` — review each one.

## Steps

0. **Dependencies.** `soffice` (LibreOffice) and the Jomolhari font:

   ```
   brew install --cask libreoffice font-jomolhari
   cp ~/Library/Fonts/Jomolhari-Regular.ttf /Applications/LibreOffice.app/Contents/Resources/fonts/truetype/
   ```

   The `cp` is needed because LibreOffice 26.8 on this Mac does **not** see user/system fonts in headless mode (only its
   bundled ones), so Tibetan renders blank otherwise. After the first render check
   `pdffonts <out.pdf> | grep -i jomolhari`. Cambria is not installed; LibreOffice substitutes Caladea (metric-compatible),
   which is fine for checking — recipients open the file in Word with Cambria.

1. **Dry run.** `python3 <skill>/scripts/parkhang.py build text.md --dry-run` prints `lineno<TAB>STYLE<TAB>text[:60]` per
   paragraph on stdout and INFO/WARN lines on stderr. Review the INFO heading lines and warnings; fix the md if needed.
   The `INFO header1/header2/footer: ...` lines show what the document will carry; set them with front matter or the
   `build` options `--header1 S --header2 S --footer S` (CLI wins over front matter; `--header2 ""` suppresses a line), e.g.
   `--header1 "DÜDŽOM TERSAR / PUDRI REKPUNG" --footer "ORGYEN KHANDRO LING"`.
   Counts per style: `... --dry-run 2>/dev/null | cut -f2 | sort | uniq -c`.

2. **Build** next to `text.md`, named after the text's folder:
   `python3 <skill>/scripts/parkhang.py build text.md -o "<folder name>.docx" [--header1 S] [--header2 S] [--footer S]`. Exit 1 = data loss (endnote marker/note
   mismatch, unreadable `cover_image`, unclosed front matter): fix the cause, never ignore it.

3. **Validate:**

   ```
   ~/.local/bin/uv run --with lxml --with defusedxml python /Users/prokop/.claude/skills/synced/bc431680-eff3-4d34-a206-4bc87dee0b24_af0f13ac-5697-4233-ad3a-8f0ac0ec1ffa/docx/scripts/office/validate.py <out.docx>
   ```

   Must print `All validations PASSED!`.

4. **Render** into a temp dir, then look at the pages:

   ```
   soffice --headless --convert-to pdf --outdir <tmpdir> <out.docx>
   pdftoppm -png -r 70 <tmpdir>/<name>.pdf <tmpdir>/p   # writes p-1.png…, zero-padded (p-01.png) from 10 pages on
   ```

   Read the PNGs (cover, first body pages, a page with cues and drum syllables, the last page).

5. **Visual checklist:**
   - Cover: title block centered, image centered below it, bottom block (Tibetan line, publisher, year) below the image
     after a 36 pt gap; no header or footer on the cover.
   - From page 2: header 2 lines flush right, footer red with the page number flush right.
   - Tibetan is in Jomolhari (not blank, not boxes) everywhere, including cue ornaments.
   - Verse = red Tibetan 19 pt / bold CAPS phonetics / regular translation, 7 pt gap after the translation.
   - Rubric italic with bold non-italic mantra words (`OM AH HUNG`).
   - Headings red caps (H1 14 pt, H2 12 pt) under black Tibetan; the body title starts a new page.
   - Cues flush right and red; drum syllables blue; `—` separators red.
   - Endnote references superscript in the text; lowercase-roman notes at the end under the notes label and Word's
     separator rule.

6. **Report to the user:** the `.docx` path, page count, warnings, count per style (from the dry run), and anything the
   `short()` heuristic decided. State that no PDF is delivered.

## Styles

Everything is a named style (sizes in pt, spacing in pt; Normal = Cambria 12, single spacing, right tab at 16 cm).

| Style | Font, size, color | Spacing before/after |
|---|---|---|
| TibVerse | Jomolhari 19, red C00000, keepNext | 0/0 |
| Phonetics | Cambria bold 12, keepNext | 0/0 |
| MantraPhonetics (based on Phonetics) | as Phonetics, keepNext off | 0/7 |
| Translation | Cambria 12 | 0/7 |
| Separator | Cambria 12, red | 7/0 |
| RubricTib | Jomolhari 14, keepNext | 0/0 |
| Rubric | Cambria italic 11 | 0/7 |
| Heading1Tib / Heading1 (`heading 1`, outline 0) | Jomolhari 18 / Cambria bold caps 14 red | 0/0 / 0/16, keepNext |
| Heading2Tib / Heading2 (`heading 2`, outline 1) | Jomolhari 14 / Cambria bold caps 12 red | 0/0 / 0/8 |
| Cue (character) | Cambria bold 10, red | – |
| Drum (character) | color blue 007BB8 | – |
| EndnoteReference (character) / EndnoteText | superscript / Cambria 11 | – / 0/7 |
| Header / Footer | Cambria bold 10, right-aligned / red | 0/0 |
| CoverTib / CoverTitle / CoverSub | Jomolhari 24 / Cambria bold caps 22 red / Cambria 14, all centered | 0/0 |
| CoverImage | centered, fit into 99 x 147 mm | 0/0 |
| CoverBottomTib / CoverBottom | Jomolhari 19 / Cambria bold 12, centered | 36/0 on the first line, else 0/0 |
| Colophon | Cambria 11 | 0/7 |
| NotesLabel | Cambria bold 12 | 14/0 |

Because everything is style-based, the user can restyle the whole document in Word through the Styles pane. Direct
formatting exists only for Tibetan runs inside Latin lines (`tib_font`), the cue tab, the page break before the body
title, and the 36 pt gap above the cover bottom block.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Tibetan blank in the render | Font not visible to LibreOffice: copy it into the bundle (step 0), or `tib_font` names a family that is not installed |
| Tibetan boxes in Word | Set `tib_font` to the family name the recipient has (e.g. `Jomolhari-ID`) and rebuild |
| `N endnote markers but M notes` | Marker not recognised (add `^` before the numeral, e.g. `slovo^v`), a stray numeral was taken as a marker, or the notes list is wrong. The message lists both sides. Build still succeeds (warning only): no endnotes, numerals stay as text, notes are plain `Colophon` paragraphs after `NotesLabel` |
| Same error plus "superscript digits found" | Old-format text; not supported |
| Heading misfire | `#` / `##` prefix to force a heading, trailing `.` to force a rubric |
| Cover bottom block not at the page foot | By design: a fixed 36 pt gap (`GAP` constant), because a bottom-anchored `framePr` lands on the next page in LibreOffice. Pin it to the bottom margin by hand in Word if wanted |
| `ERROR build: cover_image ...` | Path is relative to the md; only PNG/JPEG are accepted |

## Self-test

`python3 <skill>/scripts/parkhang.py selftest` builds `scripts/fixture/text.md` and asserts the style sequence, endnotes,
marker rules and `short()`. Run it after any change to the script.
