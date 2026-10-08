---
title: Gantt chart example
lang: en
---

# Gantt chart example

This is the shape the matching snippet inserts. It names itself, describes
itself, and sets no theme, so Check Markdown, Alt+F9, reports nothing.

```mermaid
gantt
    title Project schedule
    accTitle: Project schedule
    accDescr: Two tasks in one section, the second beginning after the first finishes.
    dateFormat YYYY-MM-DD
    section Planning
    Draft the plan :a1, 2026-09-01, 7d
    Review the plan :after a1, 5d
```
