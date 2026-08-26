---
title: What's New in EdSharp
subtitle: Conversions that work, compilers that are set up, and AI that runs on your own computer
author: Jamal Mazrui
---

# What's New in EdSharp

### Conversions that work, compilers that are set up, and AI that runs on your own computer

- [EdSharp project page on GitHub](https://github.com/JamalMazrui/EdSharp)
- [EdSharp installer for Windows](https://github.com/JamalMazrui/EdSharp/releases/latest/download/EdSharp_setup.exe)

EdSharp is a text and code editor for Windows, built for people who work by
keyboard and screen reader. It is free and open source. Earlier this year I
released a 5.0 beta, the first public release in over a decade. This is the
release that comes out of beta, and a great deal has changed since that
announcement.

## Documents convert, and keep their shape

Plain text, Markdown and HTML are now supported as both input and output,
alongside Word documents, PowerPoint files, spreadsheets, rich text, and web
pages. Open any of them with Control+Shift+O and read it as text; write any of
them with Alt+Shift+E.

PDF files were the hard case, and they now open properly. A PDF becomes
Markdown with its headings still headings, its lists still lists and its
tables still tables — so Control+B walks it by block, and the structure that a
sighted reader gets from the layout is there to navigate. Nothing about this
needs Microsoft Office.

## Spell check and a thesaurus that need nothing bought

Spell check used to require Microsoft Word. Press F7 now and EdSharp uses
Hunspell, the engine behind LibreOffice and Firefox, with a dictionary that
ships in the box. It walks the document one misspelling at a time, the way
Compile walks errors: it selects the word, spells it out, tells you where you
are in the pass, and offers suggestions you can arrow through or type over.
Words you add go into a plain file you can edit.

Shift+F7 gives synonyms from WordNet, grouped by meaning, so looking up
"light" as in weight never offers you words about illumination.

## Compilers are set up when you pick them

Control+Shift+F5 chooses a compiler, and one choice now brings everything with
it. Python runs with the official python.org build, jumps to the exact line and
column of an error, indents four spaces, and opens Python's own prompt with
your file already loaded. JavaScript runs with Node and opens Node's shell.
C# compiles with the compiler inside Windows itself — no Visual Studio, no
software development kit — and produces a 64-bit program.

Compile speech starts at the earliest error in the file rather than the first
one the tool happened to print, with your own path abbreviated away, so you
hear the problem rather than a recital of folder names.

## A console for writing snippets

Control+Shift+G opens a prompt where the editor window and its text box are
already in scope. Type an expression and see its value; type a statement and
watch it change your document. What works at the prompt works in a snippet,
which is what makes it worth having. There are two: JScript, revived from a
program I wrote in 2010, and C#, which compiles each line with the same
compiler Compile uses and loads it into the running editor.

## AI on your computer, not somebody's server

Press F12 to ask a question. If a source file is open, the question goes to a
model trained on code. If prose is open and your wording refers to it, the
document goes along; if you are just asking something, it does not. Shift+F12
always sends the document, for summarizing and rewriting.

Alt+Shift+F7 translates between eighteen languages. A better translation model
is a checkbox in the installer; tick it and translation uses it by itself.

All of it runs on your own machine. No account, no key, no limit, and no
document leaving the computer.

## Gathering, and doing the same thing to many files

Alt+7 turns the current window into a collector: everything you copy is added
to it, so three sources become one document in order. It is the quickest way
to take notes from several places, and EdSharp says so whenever you return to
such a window.

For work across many files, Alt+Shift+F finds them and puts their paths in a
window you can edit, and Alt+Equals applies a set of search and replace tasks
to every file in that list. Regular expressions run throughout: find, replace,
and two commands that pull matching text out of a document into a new one.

## Tutorials by role

There is a new set of quick starts, opened with Control+Shift+F1, written for
particular kinds of work rather than for the program in the abstract: Python
developer, NVDA add-on developer, JAWS script developer, web developer, C#
developer, translator, magazine writer, journal author, slide presenter,
document summarizer, web researcher, and batch conversion operator. Each names
the settings worth changing and the keys worth learning.

## Speech that does not repeat itself

A screen reader already announces window titles and the name, role and value
of whatever has focus. EdSharp now leaves those to it and speaks only what
only EdSharp knows — a count, a level, the text a navigation command reached.
Each message goes to one voice: JAWS if it is running, otherwise NVDA,
otherwise a system notification that Narrator reads.

Key Describer, Control+F1, has been finished properly. Every command in the
program has a description, checked automatically against the real key
bindings, so nothing answers "no description available" any more. And Alt+F4
now leaves the mode and closes the program, rather than describing itself
while you are trying to get out.

## Try it

- [EdSharp project page on GitHub](https://github.com/JamalMazrui/EdSharp)
- [EdSharp installer for Windows](https://github.com/JamalMazrui/EdSharp/releases/latest/download/EdSharp_setup.exe)

The installer's last page offers the optional pieces as checkboxes, each
saying what it will do and how large it is: Python and the document tools are
ticked, since they are what makes EdSharp itself better; Git, Node and the AI
models are not, since they serve some people and not others. Whatever you tick
installs after you press Finish, and a single summary then reports every one
by name.

I welcome feedback, by direct message if possible — email, Facebook, LinkedIn
or text. Tell me where it goes wrong.

Jamal Mazrui
