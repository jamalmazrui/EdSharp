---
title: Flowchart example
lang: en
---

# Flowchart example

This is the shape the matching snippet inserts. It names itself, describes
itself, and sets no theme, so Check Markdown, Alt+F9, reports nothing.

```mermaid
flowchart TD
    accTitle: Release decision
    accDescr: Work begins at Start, reaches a decision on whether it is ready, and either ships or goes back for fixing before trying again.
    A[Start] --> B{Is it ready?}
    B -->|Yes| C[Ship it]
    B -->|No| D[Fix it]
    D --> A
```
