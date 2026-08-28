# EdSharp

A text and code editor for Windows, built for people who work by keyboard and
screen reader. Free and open source.

- [EdSharp project page on GitHub](https://github.com/JamalMazrui/EdSharp)
- [EdSharp installer for Windows](https://github.com/JamalMazrui/EdSharp/releases/latest/download/EdSharp_setup.exe)

EdSharp edits plain text, Markdown, HTML, rich text and source code; opens
Word documents, PDF files, slide decks and spreadsheets as text you can read
and search; compiles and runs programs, landing the cursor on the error;
checks spelling without Microsoft Word; and talks to an AI model running on
your own computer for translation, summarizing and questions about code.

## Installing

Run the installer and accept the defaults. On its last page you are offered
optional pieces as checkboxes, each saying what it will do and how much space
it needs:

- **Python** and the **document tools** are ticked. Together they cost about
  a hundred and fifty megabytes and give rich PDF conversion and the
  thesaurus.
- **Git**, **Node.js** and **Ollama** with its AI models are not ticked. Each
  backs a real feature, but each serves some people and not others, and the
  models are large.

Whatever you tick installs after you press Finish, and one summary then
reports each item by name. That summary is saved in the logs folder under your
local application data, and you can see it again by running
summarizeSetup.cmd from the program folder.

EdSharp needs nothing else. It runs on the .NET Framework built into every
modern Windows, and it is a 64-bit program.

## Quick Start

Start EdSharp with Alt+Control+E from anywhere in Windows.

Type something, then try these:

- **F1** opens this guide's larger companion, the user guide.
- **Control+F1** turns on Key Describer: press any key to hear what it does
  without doing it, and press Control+F1 again to leave.
- **Alt+F10** lists every command alphabetically, with its key.
- **Control+O** opens a file; **Control+Shift+O** opens one in another format,
  converting it to text you can read.
- **F7** checks spelling. **Shift+F7** gives synonyms for the word at the
  cursor.
- **Control+F5** compiles or runs the file you are editing and takes you to
  the first error. **Control+Shift+F5** chooses which compiler to use.
- **F12** asks an AI model on your computer a question. **Alt+Shift+F7**
  translates the document or the selection.
- **Alt+Shift+E** writes the document out as something else: a Word document,
  a web page, a slide deck, plain text.

Nothing here needs a mouse, and every command has a name, a key and a
description you can hear.

## The Documentation

- **EdSharp.md** — the user guide: every command, every setting, and how the
  parts fit together.
- **Tutorials.md** — quick starts by role: Python developer, NVDA add-on
  developer, JAWS script developer, web developer, C# developer, translator,
  magazine writer, journal author, slide presenter, document summarizer, web
  researcher, batch conversion operator. Control+Shift+F1 opens it.
- **Hotkeys.md** — every command with its key, listed by command and by key.
  Generated from the program itself, so it cannot fall behind.
- **FAQ.md** — the questions people actually ask.
- **History.md** — what changed and why, newest first.
- **Development.md** — how to build EdSharp and how it is put together.
- **Announce.md** — what is new in this release.
- **License.md** — the MIT License, and the licenses of the parts that came
  from elsewhere.

Each is also here as a web page with the same name and an .htm extension,
which is what the Help menu opens.

## Reporting a Problem

EdSharp writes a log for every session, in the logs folder under your local
application data. **Control+F12** copies that log's path to the clipboard as a
file, ready to attach to a message. Sending it with a description of what you
were doing is the fastest way to a fix.

## License

EdSharp is free software under the MIT License: use it, copy it, change it,
and pass it on, including in something you sell, as long as the copyright
notice travels with it. The full terms are in License.md, and in License.htm
beside it.

Some parts of EdSharp were written by other people and keep their own
licenses. License.md names each one.

Jamal Mazrui
