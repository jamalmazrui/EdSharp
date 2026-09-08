---
title: Class diagram example
lang: en
---

# Class diagram example

This is the shape the matching snippet inserts. It names itself, describes
itself, and sets no theme, so Check Markdown, Alt+F9, reports nothing.

```mermaid
classDiagram
    accTitle: Document classes
    accDescr: A base class with one field and one method, and a derived class that adds a second method.
    class Document {
        +string title
        +open()
    }
    class Report {
        +publish()
    }
    Document <|-- Report
```
