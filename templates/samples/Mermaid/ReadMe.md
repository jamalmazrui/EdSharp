---
title: Accessible Mermaid diagrams
lang: en
---

# Accessible Mermaid diagrams

Mermaid turns text into a diagram. The text stays readable, which is the
reason it is worth using: a screen reader can read the source even when it
can make nothing of the picture.

That last part is not a nicety. Mermaid does not tell assistive technology
how its nodes connect to each other. A screen reader reaching a finished
diagram may announce the labels in an order nobody chose and say nothing
about the arrows between them. So the words you write around a diagram are
not a caption for it. They are the diagram, for some of your readers.

Five things make a Mermaid diagram usable. EdSharp's Check Markdown command,
Alt+F9, reports the first four.

## 1. Give it a name

Put a `title` in the document's front matter, as this file does. It works for
every kind of diagram and it shows to sighted readers too, which is why it is
the method to prefer. An `accTitle` line inside the diagram also names it, but
it does not display.

## 2. Describe it in full

Add an `accDescr` line saying what the diagram shows. Write it as though the
picture were missing, because for some readers it is. "A flowchart" is not a
description. "Work begins at Start, reaches a decision on whether it is
ready, and either ships or goes back for fixing before trying again" is.

## 3. Do not set a theme

An `%%{init}%%` block that sets `theme` overrides the automatic dark-mode
handling that GitHub and other sites apply, and the result is usually
unreadable in dark mode. Leave the theme alone and let the page decide.

## 4. If you set colours, check them three ways

If you do set `themeVariables`, look at the result with dark mode on, with it
off, and with Windows High Contrast. Every Mermaid theme has known contrast
problems, so a colour you choose yourself deserves the same suspicion.

## 5. Give several diagrams a heading each

A page with more than one diagram wants a heading before each, so a reader
can reach the one they want instead of arrowing through all of them.

## The examples

The files beside this one show the pattern for the diagram types EdSharp
ships snippets for. Each is short, names itself, describes itself, and sets
no theme. Open one and press Alt+F9: it should report no findings. Then
delete its `accDescr` line and press Alt+F9 again to see what the check says.

To insert any of them while writing, press Alt+V for Snippets and choose one
beginning with Mermaid. Each prompts for its name and description first,
which is the point: the accessible version is the one that costs nothing
extra.
