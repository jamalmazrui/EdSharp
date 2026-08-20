---
title: "The toolchain reissued, and three faults repaired"
subtitle: "Work done from the files you uploaded while the browser run was going"
date: 2026-08-03
lang: en
---

# What I did while the run was going

The gap I named as the one that blocks everything else was the missing tooling.
Three of those tools are now rebuilt and, more to the point, **run against your
own 137 guides** rather than shipped untested.

## auditGuides.py — measures, changes nothing

Rebuilt with the sort key the builders use: case-insensitive, ignoring a
leading "A", "An" or "The", digit runs by value, digits before letters, and
**the explicit anchor stripped before anything is compared**. Seven checks in
the file prove each of those, because both times this tool has been wrong it
was the key that was wrong, not the guides.

Two faults in my first draft, caught by running it:

- It read **every** list item shaped `- [label](#anchor)`, anywhere in a file,
  as a contents entry — so it swept up the see-also lists inside articles and
  reported ten guides as having a contents entry pointing nowhere. None of them
  did. It now reads the contents section only.
- Its change-log gate flagged **Dropbox** for publishing help about its own
  version-history feature. Whether an ambiguous title is a change log is
  decided on the body, by the removal tool, not on the title by the auditor.

## dropChangeLog.py — and nine change logs really were still there

The old rule told a change log from a feature article **by level**: an
ambiguous term like "version history" was refused only when it named a whole
category, because "Dropbox version history overview" is what a reader wants.
That was too blunt, and it left nine change logs sitting at article level.

This version **decides on the body, not the title**. A change log is a list of
versions: if an article's own text and subheadings carry three or more version
markers, it is a change log whatever it is called. Run `--report` first; it
names every candidate with its count and changes nothing.

What it found and removed:

- **Zotero — nine articles, 58,050 words.** Every release note from Zotero 1.0
  to Zotero 7, plus the live change log for 8.0. Removing them took **more than
  a third of that guide**, which is the honest measure of how much of it was
  release notes.
- **VoiceVista — one**, "Version history since the Soundscape resurrection".
- **WordPress — one**, and this one is a judgement call you may want to
  reverse. "Learn about WordPress origins and version history" opens with two
  sentences about where WordPress came from and then lists every release from
  7.0 back through the years. Under your absolute rule that is a change log
  with a preamble in front of it, so it went. Restoring it is one line.

What it correctly kept, having looked: Dropbox's two version-history articles
(a feature of Dropbox), Alexa's announcements, LinkedIn's and Instagram's
"recent changes" articles, Slack's announcement channel help.

**The safety valve fired on Zotero and was right to** — it refused to write a
file losing 36 per cent of its words. That guard exists against a buggy
removal, not against a large honest one, so it now names the figure and waits
for `--allow-large` rather than either refusing silently or writing anyway.

## fixGuides.py — heading levels and title litter

Two repairs only, and it refuses to guess at anything else.

- **Heading levels**, renumbered by depth as encountered and relative to the
  parent. Disqus and VLC each jumped h3 to h5; ten headings corrected between
  them, no word and no article changed.
- **Section markers.** Zotero's estate is a wiki, and its edit-section hash was
  left on the end of 121 headings — a screen reader reads "Citation Styles
  hash". Removed from the heading and from the matching contents entry, with
  the anchor untouched so every cross-reference still resolves.

Its valve caught a real subtlety: a bare `#` is a token, so removing 121 of them
from headings and 121 more from contents labels reads as 242 words lost and the
valve refused a correct repair. It now normalises markers on both sides before
comparing.

## Where the collection stands now

Re-audited after the repairs: **137 guides, 34,865 articles, 22,117,644
words**, and every structural gate green. Two things remain, and only one is
actionable:

- **Thirty guides carry a dropped table.** The content is gone; only a
  re-harvest brings it back.
- **Forty-eight guides have at least one category holding a single article.**
  The rule that such a category folds into Miscellaneous has always been
  applied at build time and never retroactively. It is cosmetic, it would move
  a few hundred articles around, and I did not do it without asking.

## Also ready: the Peacock builder

`buildPeacockGuide.py` is updated for the run you have going. It takes the
title from the second h1 and the category from the breadcrumb, as before —
neither of those changed. What changed is the preamble: the first edition told
readers that Peacock loads its lists on scroll, which was my wrong diagnosis.
It now writes the completeness note from what actually arrived, and says so
plainly if fewer than 90 of the sitemap's 115 articles turn up.

Send the zip and log when the run finishes and the guide can be built
immediately.

## Summary

- Three tools rebuilt and run against your own guides: `auditGuides.py`,
  `dropChangeLog.py`, `fixGuides.py`.
- Two faults in my own auditor caught by running it — a contents regex that
  read see-also lists, and a change-log gate that flagged Dropbox's feature
  article.
- Eleven change logs removed: nine from Zotero (58,050 words, over a third of
  that guide), one from VoiceVista, one from WordPress that you may want back.
- Ten heading levels corrected and 121 section markers removed.
- The collection now stands at 137 guides, 34,865 articles, 22,117,644 words,
  every structural gate green.
- `buildPeacockGuide.py` is ready for the run in flight.
