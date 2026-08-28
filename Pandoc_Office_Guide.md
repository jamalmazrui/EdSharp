# Pandoc Office Guide

How to get the most out of Pandoc with Microsoft Office file formats --
Word (.docx), PowerPoint (.pptx), and Excel (.xlsx) -- from the command
line on Windows, with EdSharp in mind. Facts come from the current
[Pandoc User's Guide](https://pandoc.org/MANUAL.html) and the
[Pandoc release notes](https://pandoc.org/releases.html); tool links are
to their project pages.

## First, check what your Pandoc can do

Pandoc grows fast, and two of the abilities in this guide are new: the
PowerPoint READER and the Excel READER were added in 2026. Check your
copy before relying on them:

- pandoc --version shows the version.
- pandoc --list-input-formats shows the readers. If pptx and xlsx are
  in that list, everything below applies. If not, run installPandoc
  from the EdSharp folder to fetch the current release.
- pandoc --list-output-formats shows the writers. There is a docx
  writer and a pptx writer; there is NO xlsx writer.

The one-line summary of Office support:

- Word: read AND write, with the richest feature set.
- PowerPoint: write (mature) and read (new).
- Excel: read (new) -- each worksheet becomes a section containing a
  table. Writing spreadsheets is not Pandoc's job.

## Word documents (.docx)

### Reading Word files

The basic command reads a Word file into any other format:

    pandoc report.docx -f docx -t markdown -o report.md

Options and extensions that change what you get:

- --extract-media=FOLDER pulls the pictures and other media out of the
  .docx into a folder and rewrites the references to point at them.
  Without it, images are dropped from plain-text-like outputs.
- --track-changes=accept is the default: insertions and deletions from
  Word's Track Changes are applied. reject ignores them. all keeps
  everything, wrapped in labeled spans with author and time -- useful
  for scripts that accept only one reviewer's changes.
- -f docx+styles keeps EVERY Word style name as an attribute on the
  text it covered, even styles Pandoc does not understand. Use it when
  the style names themselves carry meaning you must not lose.
- -f docx+citations turns citations inserted by the Zotero, Mendeley,
  or EndNote Word plugins into real Pandoc citations instead of plain
  text -- the key to rescuing a manuscript back out of Word without
  losing its bibliography.

### Writing Word files

The basic command:

    pandoc article.md -f markdown -t docx -o article.docx

The look of the output is controlled by a REFERENCE DOCUMENT, not by
templates. Pandoc ignores the reference file's contents and uses only
its styles, margins, page size, headers, and footers:

1. Make your starting copy:
   pandoc -o custom-reference.docx --print-default-data-file reference.docx
2. Open custom-reference.docx in Word, change the styles (Normal, the
   Heading levels, Title, Author, Abstract, Block Text, Source Code,
   Footnote Text, Table, and so on), and save. Change styles only; the
   text in the file does not matter.
3. Use it: add --reference-doc custom-reference.docx to the command.
   If a file named reference.docx sits in Pandoc's user data folder,
   it is used automatically.

Two more writing powers:

- Custom styles: wrap text in a fenced Div or a Span with a
  custom-style attribute, and the output uses that Word style by name.
  For example: ::: {custom-style="Warning"} ... ::: gives every
  paragraph inside the Word style called Warning.
- -t docx+native_numbering numbers figures and tables with Word's own
  counter fields, so they renumber themselves when edited later.

## PowerPoint slide shows (.pptx)

### Writing slides

    pandoc talk.md -f markdown -t pptx -o talk.pptx

- Headings at the SLIDE LEVEL start new slides; headings above it make
  section-divider slides; headings below it are subheads inside a
  slide. Pandoc picks the slide level from the document's shape, or
  --slide-level=NUMBER sets it; level 0 means slides split only at
  horizontal rules.
- --incremental makes list items appear one by one.
- Speaker notes: a Div marked notes becomes PowerPoint presenter
  notes: ::: notes ... :::
- Design comes from a reference file, exactly as with Word:
  pandoc -o custom-reference.pptx --print-default-data-file reference.pptx
  then edit its slide masters and layouts in PowerPoint and pass it
  with --reference-doc. Pandoc fills the layouts by their standard
  names, so keep the layout names.

### Reading slides

The new pptx reader turns a slide show into an outline of its text:

    pandoc talk.pptx -f pptx -t markdown -o talk.md

It is text-focused: expect the headings, bullets, and table text, not
the visual design. Combine with --extract-media to save the pictures.

## Excel workbooks (.xlsx)

The new xlsx reader turns each worksheet into a section containing a
table, which then converts to anything Pandoc writes:

    pandoc figures.xlsx -f xlsx -t gfm -o figures.md
    pandoc figures.xlsx -f xlsx -t docx -o figures.docx

Notes:

- This reads the VALUES of a workbook as tables. Formulas, charts, and
  formatting are not what it is for.
- There is no xlsx writer. To CREATE a spreadsheet, leave Pandoc: in
  EdSharp's world, the OfficeConvert utilities and 2htm cover the
  directions Pandoc does not.
- Pandoc also reads plain CSV and TSV files as tables (-f csv, -f
  tsv), which remains the simplest bridge for raw data.

## The command-line options that matter most for Office work

- -f and -t name the input and output formats; add +extension or
  -extension to a format name to turn a feature on or off.
- --reference-doc styles docx and pptx output.
- --extract-media saves embedded pictures when reading docx or pptx.
- --track-changes controls Word revision marks when reading docx.
- --citeproc, --bibliography, and --csl handle citations (see the
  tutorial below).
- --data-dir names Pandoc's user data folder, where a default
  reference.docx and a csl folder of styles can live.
- -d FILE (a defaults file) stores a whole set of options in one YAML
  file, so a long command becomes: pandoc -d journal article.md

## Helpers from GitHub worth knowing

- [pandoc-crossref](https://github.com/lierdakil/pandoc-crossref)
  numbers figures, equations, and tables and resolves references to
  them, in docx among other outputs. Use a build matched to your
  Pandoc version; the project page says mismatches misbehave.
- [pandoc-zotxt.lua](https://github.com/odkr/pandoc-zotxt.lua) looks
  up citation keys directly in a running Zotero, so you cite without
  exporting a bibliography file first.
- [pantable](https://github.com/ickc/pantable) renders CSV data as
  proper Pandoc tables inside your document.
- The [Pandoc filters list](https://github.com/jgm/pandoc/wiki/Pandoc-Filters)
  on the Pandoc wiki catalogs many more, including docx-specific
  helpers.
- Outside the filter world, the npm tool
  [xlsx2md](https://www.npmjs.com/package/xlsx2md) converts workbooks
  to Pandoc grid tables with formatting niceties the built-in reader
  does not attempt.

## How EdSharp wires this in

EdSharp's Open Other Format (Control+Shift+O) and Export Format
(Alt+Shift+E) run the entries in the Import and Export tables of
EdSharp.inix, most of which call pandoc.exe in Convert\Pandoc
directly. With a current Pandoc, two simplifications become possible
and are worth adopting: pptx and xlsx imports can call Pandoc's new
readers directly instead of going through 2htm or a CSV bridge. Check
with --list-input-formats first, then the Import lines become, for
example:

    xlsx2md=%ProgDir%\Convert\Pandoc\pandoc.exe "%SourceLong%" -f xlsx -t gfm -o %Target%
    pptx2md=%ProgDir%\Convert\Pandoc\pandoc.exe "%SourceLong%" -f pptx -t gfm -o %Target%

Personal additions belong in the EdSharp.inix in your data folder,
which outranks the shipped one and survives upgrades.

## Tutorial: submitting a journal article, citations included

The goal: write journal_article.md in EdSharp, keep the sources in one
small text file, and produce journal_article.docx with the citations
and bibliography formatted in a named style -- using nothing but
Pandoc on the Windows command line.

### Step 1. Keep your sources in a .bib file

BibTeX is the simplest container for references: one entry per source,
readable and editable in EdSharp itself. The sample
journal_article.bib beside this guide shows an article, a book, and a
chapter. Each entry's first line holds its KEY (for example,
mazrui2024) -- that key is how you cite. Where do entries come from?

- Type them: the sample shows every field you need.
- Most journal sites and Google Scholar offer a BibTeX export for any
  paper -- paste the entry into your .bib file.
- If your library lives in [Zotero](https://www.zotero.org), the
  Better BibTeX add-on keeps a .bib file automatically exported and
  always current; Pandoc just reads that file. Pandoc also reads RIS
  and EndNote XML bibliographies directly if that is what you have.
- [JabRef](https://www.jabref.org) is a free Windows editor for .bib
  files when they grow large.

### Step 2. Cite by key in the article

In the text, a citation is a bracketed key: [@mazrui2024] renders as a
parenthetical citation; [@mazrui2024, p. 12] adds a page;
[@mazrui2024; @doe2023] cites two works; @mazrui2024 alone gives a
narrative citation (the name in the sentence). The bibliography is
appended automatically where a Div with identifier refs sits, or at
the end. The sample article demonstrates each pattern.

### Step 3. Convert with citations turned on

Chicago author-date is Pandoc's BUILT-IN default style, so the whole
Chicago pipeline is one command with nothing extra installed:

    Convert\Pandoc\pandoc.exe journal_article.md -f markdown -t docx --citeproc --bibliography journal_article.bib -o journal_article.docx

For APA (or any of thousands of journal styles), add a CSL style file
once and name it with --csl:

    Convert\Pandoc\pandoc.exe journal_article.md -f markdown -t docx --citeproc --bibliography journal_article.bib --csl apa.csl -o journal_article.docx

Styles come from the CSL project's styles repository on GitHub (the
project is named citation-style-language/styles; every journal style
lives there as one .csl file). Pandoc can also fetch a style straight
from a web address given to --csl, and styles dropped into the csl
folder of Pandoc's data directory are found by bare name. The sample
article names its bibliography in its metadata block, so the command
shrinks further: pandoc journal_article.md --citeproc -t docx -o
journal_article.docx

### Step 4. Dress it for the journal

- Build a reference document from the journal's specimen: set Normal
  to the required font and spacing, the heading styles to theirs, and
  pass --reference-doc journal-reference.docx.
- link-citations: true in the metadata makes each citation a live link
  to its bibliography entry -- reviewers like it, and it costs
  nothing.
- If the journal wants a blinded manuscript, keep author metadata out
  and let the title page travel separately.

### Step 5. The round trip

When the revised manuscript comes back as a .docx full of tracked
changes and plugin citations, Pandoc brings it home:

    pandoc revised.docx -f docx+citations --track-changes=all -t markdown -o revised.md

Zotero-style citations become keys again, and every reviewer change is
labeled with its author -- ready for EdSharp.

### Wiring the tutorial into EdSharp

One line in the data-folder EdSharp.inix gives Export Format a
citation-aware docx target:

    md2docxrefs=%ProgDir%\Convert\Pandoc\pandoc.exe "%SourceLong%" -f markdown -t docx --citeproc -o %Target%

With the bibliography named in the article's own metadata, Alt+Shift+E
then produces the submission file in one stroke.
