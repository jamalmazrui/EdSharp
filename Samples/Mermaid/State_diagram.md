---
title: State diagram example
lang: en
---

# State diagram example

This is the shape the matching snippet inserts. It names itself, describes
itself, and sets no theme, so Check Markdown, Alt+F9, reports nothing.

```mermaid
stateDiagram-v2
    accTitle: Order states
    accDescr: An order starts as a draft, is submitted for review, and is then either approved or returned to draft.
    [*] --> Draft
    Draft --> Submitted
    Submitted --> Approved
    Submitted --> Draft
    Approved --> [*]
```
