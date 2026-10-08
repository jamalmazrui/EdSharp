---
title: Entity relationship diagram example
lang: en
---

# Entity relationship diagram example

This is the shape the matching snippet inserts. It names itself, describes
itself, and sets no theme, so Check Markdown, Alt+F9, reports nothing.

```mermaid
erDiagram
    accTitle: Customers and orders
    accDescr: One customer places many orders, and each order contains many items.
    CUSTOMER ||--o{ ORDER : places
    ORDER ||--|{ ITEM : contains
```
