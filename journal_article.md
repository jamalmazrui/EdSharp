---
title: "Screen Reader Productivity in Text Editors: A Command-Line Approach"
author: "Jamal Mazrui"
abstract: |
  This sample article shows every citation pattern the tutorial in
  Pandoc_Office_Guide.md describes. Replace its text with a real
  manuscript and the same commands keep working.
keywords: [accessibility, screen readers, text editors]
bibliography: journal_article.bib
link-citations: true
---

<!-- Build it (Chicago author-date, Pandoc's built-in default):
Convert\Pandoc\pandoc.exe journal_article.md --citeproc -t docx -o journal_article.docx
For APA, add: --csl apa.csl  (one .csl file from the CSL styles repository)
-->

# Introduction

Editing efficiency for screen reader users depends less on visual
layout than on predictable structure [@mazrui2024]. Earlier work on
audio interfaces reached a similar conclusion from a different
direction [@doe2023, pp. 41-42]. As @chen2022 argues, the command line
remains the most scriptable environment for repeatable document work.

# Method

We measured task completion for common editing operations, following
the protocol of @doe2023 with the corrections suggested by later
replications [@mazrui2024; @chen2022, chap. 3].

# Results

Structured hotkeys outperformed menu navigation on every task. The
effect was strongest for conversion tasks, where a single keystroke
replaced a multi-step dialog.

# Discussion

The results support building conversion into the editor as first-class
commands. Configuration through layered plain-text files keeps such
commands correctable by the very population that depends on them
[@mazrui2024].

# References

::: {#refs}
:::
