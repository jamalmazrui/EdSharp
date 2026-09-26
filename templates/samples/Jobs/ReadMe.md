---
title: Transform job files
lang: en
---

# Transform job files

Transform Files works through a job: a text file with one section per task,
each saying what to find and what to do with it. It runs over as many files
as you choose, so a change you would make by hand a hundred times is made
once.

Each section may hold:

- **Find** — a regular expression. The only required key.
- **Replace** — what to put in its place. Leave it out to delete matches.
- **Options** — Multiline, IgnoreCase, Singleline, and the rest.
- **Extract** — set to 1 to gather matches to the clipboard instead of
  changing the file.
- **Divider** — what to put between gathered matches. A form feed and a
  newline unless you say otherwise.

## Three things that are easy to miss

**A space at either end needs quoting.** A value on one line is trimmed, so
`Replace= x` gives you `x` and the space is gone. Put quotation marks round
it and both characters survive:

    Replace=" x"

The quotation marks are the wrapper and do not appear in the result. This is
the settings format's own way of protecting spaces, and it works for the
Find value too.

**$# is the match number.** In a Replace value it becomes 1 for the first
match, 2 for the second, and so on. That is how you number a series of
things that all look alike — see NumberPassages.inix.

**$1 is the first captured group.** Anything you put in brackets in the Find
comes back as $1, $2 and so on, so a page number you matched can be kept.

## The examples

- **NumberPassages.inix** — passages sharing an identical ending, numbered
  in order.
- **BrailleToText.inix** — a braille-ready file turned into something to
  read: page indicators kept in a findable form, form feeds and runs of
  blank lines removed. Shows both the quoting and $1.
- **ExtractHeadings.inix** — every Markdown heading to the clipboard,
  changing nothing.

## Trying one safely

Transform Files reports the matches before it changes anything. Run a job,
read the counts, and only then let it apply. A job that reports no matches
has a Find that does not fit your text, which is quicker to discover on a
report than in a file you have already rewritten.
