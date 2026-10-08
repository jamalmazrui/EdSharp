---
title: Sequence diagram example
lang: en
---

# Sequence diagram example

This is the shape the matching snippet inserts. It names itself, describes
itself, and sets no theme, so Check Markdown, Alt+F9, reports nothing.

```mermaid
sequenceDiagram
    accTitle: Request and response
    accDescr: The user sends a request to the program, and the program sends a response back.
    participant A as User
    participant B as Program
    A->>B: Request
    B-->>A: Response
```
