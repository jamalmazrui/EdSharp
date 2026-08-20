---
title: "OSARA built, and the wiki harvester you approved"
subtitle: "2 August 2026"
date: 2026-08-02
lang: en
---

# OSARA_Help_Guide.md

**3 categories, 25 articles, about 10,100 words.** All eight gates green.

The harvest found exactly one Markdown document in the repository -- `readme.md`,
11,207 words, with 9 second-level sections, 24 third-level topics and 34 below
that. So the readme really is the whole of OSARA's own documentation, and the
cross-check agreed: 10,396 words rendered against 11,207 in the source, the
difference being the navigation furniture the rendered page carries.

The arrangement is the author's own. Usage becomes a category holding its
eighteen topics, Download and Installation holds Windows, Mac and Key Map, and
the four sections with nothing beneath them -- Requirements, Additional
Documentation, Support, Reporting Issues -- become articles in Miscellaneous.

Set aside and named in the preamble: **Building**, **Translating** and
**Contributors**. Those are written for people working on OSARA, not people
using it, which is the same rule that excludes developer documentation
everywhere else in the collection.

Because the source was already Markdown, nothing passed through a conversion
that could alter it. That is the first volume in the collection where that is
true.

# The wiki, now that you have approved it

`getReaperAccessibilityHelp.py` gathers reaperaccessibility.com.

It runs MediaWiki, so the estate is enumerated through the wiki's own API rather
than guessed at -- the same move that rescued Libsyn and Brave and that scopes
the VLC script. The installation's API address varies, so both usual locations
are tried and the log says which answered. Coverage is stated at the end: the
wiki's own article count beside the number saved.

Refused: talk pages, user pages, edit and history views, old revisions, Special
pages, files and templates, and anything outside the main namespace.

**One thing is hard-coded on purpose.** The function that decides whether a page
is the publisher's own writing returns false for every page here, always.
Nothing on that wiki is OSARA's own, and the build must not be able to drift
into presenting it as though it were.

15 predicates, and one of them earned its keep: it caught a real bug. My
adaptation dropped the `/wiki/` segment from canonical addresses, so every page
would have been requested at an address that answers 404 -- a harvest that
reports refusals cleanly and saves nothing, exactly the Brave shape. Fixed, and
all 15 pass.

# What the combined edition will say

When the wiki harvest arrives I will build one volume with both, and its
preamble will state plainly:

- Which material is OSARA's own and which is community-written, marked at the
  category level so it is never ambiguous while reading.
- That the wiki is included because OSARA's own documentation names it as where
  the in-depth guidance lives -- the reason for the exception, in the guide
  itself rather than only in our notes.
- That the community material carries no guarantee from OSARA's authors.

If you would rather ship the OSARA-only guide as it stands and keep the wiki as
a separate volume, that also works and is one command either way. The file in
this zip is complete and passes every gate on its own.

# Summary

- OSARA guide built and shipped: 3 categories, 25 articles, ~10,100 words.
- Set aside: Building, Translating, Contributors -- contributor material, named
  in the preamble.
- `getReaperAccessibilityHelp.py` delivered for the wiki, enumerated through the
  wiki's own API, with the community nature stated in the log and unable to be
  presented as OSARA's own.
- A predicate caught a real bug in it before you ran it: dropped `/wiki/`
  segment, which would have produced a clean-looking harvest of nothing.
- Run it in a fresh folder and send the zip and the log.
