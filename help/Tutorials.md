# EdSharp Tutorials

Short, practical starts for the kinds of work people do in EdSharp. Each
one names the settings worth changing, the commands worth learning, and
the keys that invoke them. Read the one that matches what you are doing;
they do not depend on each other.

Two things are worth knowing before any of them.

Settings live in Configuration Options, Alt+Shift+C. Anything named
below in capitals, such as IndentUnit, is a setting you will find there.

Every command has a name and a key. Press Control+F1 to turn on Key
Describer, then press any key to hear what it does without doing it;
press Control+F1 again to leave. Alt+F10 lists every command
alphabetically, and Alt+Shift+H shows the whole table of names, keys and
descriptions in a window you can search.

## Contents

- [Python Developer](#python-developer)
- [NVDA Add-on Developer](#nvda-add-on-developer)
- [JAWS Script Developer](#jaws-script-developer)
- [Node.js and Web Developer](#nodejs-and-web-developer)
- [C# Developer](#c-developer)
- [Language Translator](#language-translator)
- [Magazine Article Author](#magazine-article-author)
- [Journal Article Author](#journal-article-author)
- [Slide Presenter](#slide-presenter)
- [Document Summarizer](#document-summarizer)
- [Web Researcher](#web-researcher)
- [Batch Conversion Operator](#batch-conversion-operator)

## Python Developer

Python is the language EdSharp supports most fully, because Python's
whitespace is the hardest thing about writing code with a screen reader
and EdSharp answers it from several directions at once.

### Setup

Press Control+Shift+F5, Pick Compiler, and choose Python. That one
choice brings a working set of settings: Control+F5 runs your file with
the official Python from python.org and jumps to the line and column of
any error; Control+Shift+G opens Python's own prompt with your file
already run, so the functions you just wrote are there to call;
IndentUnit becomes four spaces, which is what PEP 8 asks and what
collaborators expect; and the comment prefix becomes the number sign, so
indentation commands know to skip comments.

Your own indentation still wins. EdSharp reads the indentation a file
already uses and follows it, so a file written with tabs stays tabbed no
matter what the compiler setting says. The setting only governs a file
with no indentation yet.

For wxPython, install the library once from a command prompt:

    python -m pip install wxPython

Then open Samples\fruitBasket.py from the EdSharp program folder, which
is a complete wxPython program written in the Camel Type style, and
press Control+F5 to run it. Control+Shift+F2, Sample Programs, lists it
and the others.

### Working with Indentation

Four commands make Python's structure audible.

PyBrace, Alt+Shift+LeftBracket, rewrites the open document into a flat
form where structure is spoken text rather than counted spaces: a
compound statement ends with a brace, and every block closes with a line
naming the keyword it closes, such as "brace end def". Edit in that form
if you find it easier. PyDent, Alt+LeftBracket, turns it back into
indented Python, rebuilding the indentation and leaving a comment at the
end of each block so the ends of functions and loops stay audible while
you read. Dictionary literals, docstrings and line continuations survive
the round trip untouched.

Indent Mode, Alt+Shift+I, announces each change in indentation as you
arrow through code: "in 1" when a line is one level deeper than the last,
"out 2" when it is two levels shallower. It also swaps what Enter and
Shift+Enter do, so Enter keeps the current indentation.

Indentation, Alt+I, in the Query menu, says how many levels deep the
current line is.
Press it twice and it reads the whole chain of enclosing blocks,
outermost first, such as "class Greeter, def greet, if loud, for i in
range 3" -- which is the answer to "where am I?" that sighted readers get
from the layout at a glance.

Next Indent and Prior Indent, Control+I and Control+Shift+I, move to the
next and previous change of indentation. Next Block and Prior Block,
Control+B and Control+Shift+B, move by whole blocks.

Two more commands round it out. Format Code, Alt+F8, normalizes a
Python file's indentation to the unit in force. Infer Indent, Alt+RightBracket, captures the indentation of the file in
front of you into the setting.

### Key Reminders

- Control+Shift+F5 Pick Compiler, Control+F5 Compile, Control+Shift+G Go
  to Environment
- Alt+Shift+LeftBracket PyBrace, Alt+LeftBracket PyDent
- Alt+Shift+I Indent Mode, Alt+I Indentation (press twice for the chain)
- Control+I and Control+Shift+I by indentation, Control+B and
  Control+Shift+B by block
- F12 Chat with AI: with a source file open, questions go to the coding
  model when it is installed

## NVDA Add-on Developer

An NVDA add-on is Python, so everything in the Python tutorial applies.
What differs is the shape of the project and the testing loop.

### Setup

Pick the Python compiler as above. NVDA's own code follows PEP 8 with
four space indentation, so leave IndentUnit as the compiler sets it.

An add-on is a folder of Python files plus a manifest, packaged as a zip
with the .nvda-addon extension. EdSharp edits the parts; NVDA installs
the whole. Keep a build script in the project and run it with Prompt
Command, Alt+F5, which runs any command line and shows its output the
same way Compile does, jumping to errors.

Set GoToEnvironment to NVDA's own Python console if you use it, or leave
it as the compiler's prompt for testing plain functions.

EdSharp itself ships an add-on, EdSharp.nvda-addon in the program
folder, which is a working example of the layout.

### Key Reminders

- Alt+F5 Prompt Command for the build script; Control+F5 Compile for a
  single file
- Alt+Shift+F9 Run Code Blocks if your notes hold snippets to try
- Control+F12 Copy Log when reporting a problem: it puts this session's
  log on the clipboard as a file, ready to attach to a message

## JAWS Script Developer

### Setup

Press Control+Shift+F5 and choose JAWS Script. Control+F5 then compiles
the open script with scompile.exe, the compiler that comes with JAWS;
EdSharp puts the running JAWS folder on the path at startup, so whatever
version you have is the one used. Errors report a line and column and
the cursor lands there. The comment prefix becomes the semicolon and
IndentUnit becomes one tab, which is what the JAWS script files
themselves use.

JAWS script has no interactive prompt, so Control+Shift+G says so
rather than opening something unrelated.

Snippets help here more than anywhere. Alt+S saves the selection as a
snippet under the current compiler, so a script skeleton or a common
call is a keystroke away; Alt+V inserts one. EdSharp ships two, InputBox
and SayString, which appear in the list marked as shipped.

### Key Reminders

- Control+Shift+F5 Pick Compiler, Control+F5 Compile
- Alt+S Save Snippet, Alt+V Invoke Snippet
- Alt+PageDown and Alt+PageUp move between scripts and functions
- Control+Shift+G opens the snippet console when the JScript .NET
  compiler is chosen: a JScript prompt with the editor window as frm and
  its text box as rtb, so a line that works there works in a snippet

## Node.js and Web Developer

### Setup

Press Control+Shift+F5 and choose JavaScript. Control+F5 then runs the
open file with Node, and the error jump understands both forms Node
reports: the plain line for a syntax error and the line with column
inside a stack frame. Node's own frames are removed from what is spoken,
so you hear your error rather than Node's internals.
Control+Shift+G opens Node's read-evaluate-print shell. IndentUnit
becomes two spaces, the prevailing JavaScript convention, and the
comment prefix becomes two slashes.

For web pages, Preview Markdown in Web Browser and the HTML commands
matter more than the compiler. Control+H formats HTML; Control+F9
previews the current Markdown in a window, and the browser preview draws
Mermaid diagrams properly. Check Markdown, Alt+F9, reports images
without alternative text, heading jumps and bare web addresses --
accessibility faults you would otherwise ship.

Samples\Web in the program folder holds the fruit basket program written
four ways for the browser, with a Node script that serves them:

    node serveFruitBasket.js

### Key Reminders

- Control+Shift+F5 Pick Compiler, Control+F5 Compile, Control+Shift+G
  Node shell
- Alt+F9 Check Markdown, Control+F9 Preview Markdown
- Control+Shift+F2 Sample Programs lists the web samples

## C# Developer

### Setup

Press Control+Shift+F5 and choose C#. Control+F5 then compiles the open
file with the C# compiler that ships inside Windows itself -- no Visual
Studio, no SDK -- producing a 64-bit program, and jumps to the exact
line and column of the first error in the file. IndentUnit becomes four
spaces, the Microsoft convention, and the comment prefix two slashes.
Windows ships no C# console, so Control+Shift+G opens EdSharp's own
JScript interpreter for trying expressions.

Two things are worth knowing when writing Windows Forms code. First,
Samples\fruitBasket.cs is a complete Windows Forms program in Camel Type
with a build line in its comments; open it with Control+Shift+F2 and
press Control+F5. Second, when you ask F12 for help with C#, say
"targeting .NET Framework 4.8" in your question. Most C# written today
targets the newer .NET, and a model that assumes it will hand you code
that does not compile here.

### Key Reminders

- Control+Shift+F5 Pick Compiler, Control+F5 Compile
- Compile jumps to the earliest error in the file, not the first one the
  compiler printed; fix it and press Control+F5 again
- Alt+F8 Format Code runs the C family through astyle
- Control+Shift+G opens the C# console: a prompt where frm is the editor
  window and rtb its text box, so "rtb.SelectedText.ToUpper()" prints the
  answer and "rtb.SelectedText = rtb.SelectedText.ToUpper();" changes the
  document. Each line is compiled, so expect about a second per line; a
  line ending in a semicolon is kept for the rest of the session
- F12 Chat with AI uses the coding model for .cs files when it is
  installed

## Language Translator

### Setup

Translation needs Ollama and a model, both offered as checkboxes on the
last page of the installer. Tick Ollama, which brings llama3.2, and tick
qwen2.5:7b, the translation model; translation then uses the better one
by itself. Everything runs on your computer, so there is no account, no
limit and no document leaving the machine.

Translate Language, Alt+Shift+F7, translates the selection, or the whole
document when nothing is selected. One dialog asks both languages, with
your last pair already selected, and the translation opens in a new
window so the original is untouched.

Quality is good between the widely written languages -- Spanish, French,
German, Portuguese, Italian -- and weaker for languages with less
material behind them. A long document takes minutes, and EdSharp speaks
a count every fifteen seconds so you know it is working.

TranslateModel names a different model if you install one.

### Key Reminders

- Alt+Shift+F7 Translate Language
- Select first to translate a passage; select nothing for the whole
  document
- F12 Chat with AI answers questions about a translation, such as asking
  for a more formal wording

## Magazine Article Author

### Setup

Write in Markdown. Set the default extension for new documents to md if
you write mostly articles, and leave word wrap on with Control+W.

Three commands carry an article from draft to delivery. Check Markdown,
Alt+F9, reports the faults an editor will send back: headings that skip
a level, images without alternative text, bare web addresses, unclosed
code fences. Preview Markdown, Control+F9, shows it formatted;
Control+Shift+F9 opens it in your browser. Export Format, Alt+Shift+E,
writes the article as a Word document, HTML, or whatever the magazine
wants, using Pandoc.

For style work, F12 Chat with AI is the strongest tool in the program.
Ask it to tighten a paragraph, to rewrite at a ninth grade reading
level, or to suggest a title; the answer opens in its own window so you
can compare. Press F7 to spell check before sending, and Shift+F7 on a
word for synonyms grouped by meaning.

### Key Reminders

- Alt+F9 Check Markdown, Control+F9 Preview, Alt+Shift+E Export Format
- F7 Spell Check, Shift+F7 Thesaurus
- F12 Chat with AI on a selection: "make this tighter", "ninth grade
  reading level"

## Journal Article Author

### Setup

A journal article is a magazine article with citations and a required
format, so start from the section above and add three things.

Keep your references in a .bib file beside the article and name it in
the document's own metadata block at the top:

    ---
    title: Your Title
    author: Your Name
    bibliography: references.bib
    csl: apa.csl
    ---

Cite in the text with a key in square brackets, such as [@smith2020].
When you export with Alt+Shift+E, Pandoc formats the citations and
builds the reference list in the style the .csl file names. Download the
style your journal requires from the Zotero style repository and keep it
beside the article.

Tables and figures deserve attention, since they are where accessibility
is usually lost. Give every figure alternative text in the Markdown, and
give every table a caption; Check Markdown, Alt+F9, reports missing
alternative text before a reviewer does.

If the journal supplies a Word template, name it as the reference
document so your export matches their layout.

### Key Reminders

- Alt+Shift+E Export Format for the submission file
- Alt+F9 Check Markdown before every submission
- Control+Shift+O Open Other Format converts a colleague's Word document
  or a PDF into Markdown you can edit, keeping headings, lists and
  tables

## Slide Presenter

Pandoc turns Markdown into a PowerPoint file, speaker notes included, so
a talk can be written as text and never touched with a mouse.

### Setup

Write the talk in Markdown. The heading level that begins a slide is
called the slide level: with the usual arrangement, each second level
heading starts a slide and first level headings become section dividers.
A horizontal rule, three hyphens on their own line, starts a slide
without a title.

Speaker notes go in a div marked as notes, which PowerPoint shows in
Presenter View and in handouts but never on the slide:

    ## What EdSharp Does

    - Edits any text
    - Converts documents
    - Talks to a local AI

    ::: notes
    Mention that everything runs offline. Ask how many people use JAWS.
    :::

Export with Alt+Shift+E and choose pptx. To match a house style, make a
copy of your organization's template and name it as the reference
document; Pandoc uses its theme, its fonts and its layouts, choosing a
layout per slide by what the slide contains -- a title slide, a section
header, a two-content slide when you use columns, and a plain title and
content slide otherwise.

Two accessibility points worth minding, since they are yours to get
right and nobody else's: give every slide a unique title, because screen
readers and PowerPoint's own outline use titles to navigate, and give
every image alternative text in the Markdown.

You can also go the other way. Control+Shift+O opens an existing
PowerPoint file as Markdown, which is the quickest way to read a deck
somebody sent you: the titles become headings and the bullets become
lists.

### Key Reminders

- Alt+Shift+E Export Format, choosing pptx
- Control+Shift+O Open Other Format to read a deck as Markdown
- Second level headings start slides; three hyphens start an untitled
  one; a notes div holds speaker notes

## Document Summarizer

### Setup

Tick Ollama on the installer's last page so F12 has a model to talk to.

The loop is short. Open the document -- Control+Shift+O converts a PDF,
Word file, slide deck or spreadsheet into Markdown with its headings and
tables intact. Press F12, Chat with AI, and type an instruction such as
"summarize in five bullet points" or "list the recommendations only".
The summary opens in a new window, leaving the original untouched.

Two details make it work better. Select a section first if you want that
section summarized; with nothing selected the whole document travels.
And ask for a number of bullet points or paragraphs rather than a number
of words, because a model counts words poorly and structure well.

When the instruction is a general question rather than a request about
the document, F12 notices and sends the question alone, which answers in
seconds. Shift+F12, Chat about Document, forces the document to travel
whatever the wording.

### Key Reminders

- Control+Shift+O Open Other Format, F12 Chat with AI, Shift+F12 Chat
  about Document
- The status line says which context was used: with selection, with
  document, or question only
- A long document takes minutes; a count is spoken every fifteen seconds

## Web Researcher

### Setup

Research is gathering, and EdSharp is built for gathering: fetch a page,
read a PDF as text, and collect fragments from several sources into one
document without ever leaving the keyboard.

Web Download, Alt+Shift+W, picks files to download from a web page or
from the addresses in the current document. Fetch an article as HTML and press Control+Shift+O,
Open Other Format, to read it as Markdown with its headings and lists
intact -- far quieter than a browser, and searchable with Control+F.

PDFs are the usual currency of research, and they open the same way:
Control+Shift+O converts a PDF into Markdown with headings, lists and
tables preserved, so Control+B walks it by block and Alt+F9 reports
anything malformed. Tick the document tools on the installer's last page
so PDF conversion is available.

The gathering itself is the clipboard. Alt+7, Append from Clipboard,
turns the current window into a collector: everything you copy is added
to it, so you can read three sources, copy a paragraph from each, and
find them stacked in one document in order. EdSharp says "Append from
clipboard" whenever focus returns to such a window, so the mode is never
a surprise. Copy Line and Append to Clipboard, in the Edit menu, add a
single line without leaving what you are reading.

When the reading is done, F12, Chat with AI, summarizes what you
collected, and Shift+F12 asks about the whole document. Alt+Shift+F7
translates a source that arrived in another language.

### Key Reminders

- Alt+Shift+W Web Download to fetch, Control+Shift+O Open Other Format
  to read it as Markdown
- Alt+7 Append from Clipboard to collect fragments into one window
- Control+F Forward Find within a source; Control+B and Control+Shift+B
  by block
- F12 Chat with AI to summarize what you gathered

## Batch Conversion Operator

### Setup

Two commands do bulk work, and they pair.

File Find, Alt+Shift+F, searches folders and puts the matching file
paths into a window, one per line. That list is the input to everything
below, and you can edit it by hand -- delete the lines you do not want
converted.

Transform Files, Alt+Equals, applies a set of search and replace tasks
to every file named in the current window. Write the tasks first, run
them against a copy of the files first, and keep the task list as a
document you can reuse.

For format conversion in bulk, put the conversions in a script and run
it with Prompt Command, Alt+F5, whose output appears the same way a
compiler's does. EdSharp's own Convert folder holds the scripts it uses,
which are worth reading as models: each takes a source and a target and
reports what happened.

Text Convert, Control+T, converts the open document between formats one
at a time, and Export Format, Alt+Shift+E, writes it out as something
else. Control+Shift+O converts on the way in.

Every job leaves a record. The session log in the logs folder under your
local application data names every command EdSharp ran and its exit
code, and Control+F12 copies its path to the clipboard.

### Key Reminders

- Alt+Shift+F File Find to build the list, Alt+Equals Transform Files to
  work through it
- Alt+F5 Prompt Command for a script, Control+F5 Compile for a single
  file
- Control+F12 Copy Log when something needs explaining

## Using EdSharp: a hands-on walk

This walk was adapted, with thanks, from the *Using TextPal* tutorial
written by Jim Homme for TextPal, a predecessor of EdSharp that gave it
many of its ideas. It was a separate document until September 2026.

A hands-on tutorial for getting started with the EdSharp text editor.

This tutorial is adapted, with thanks, from the *Using TextPal* tutorial originally written by Jim Homme for TextPal, a predecessor of EdSharp that contributed many of its ideas. The text here has been revised to match EdSharp's terminology, concepts, and key bindings. For the complete reference, see the EdSharp User Guide (press F1 in EdSharp).

### Contents

- [Introduction](#introduction)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Common Hot Keys](#common-hot-keys)
    - [Working With Files on Disk](#working-with-files-on-disk)
    - [Navigating in the Current File](#navigating-in-the-current-file)
    - [Using the Clipboard](#using-the-clipboard)
    - [Changing Character Case](#changing-character-case)
    - [Searching and Replacing](#searching-and-replacing)
    - [Working With Sections](#working-with-sections)
    - [Spell Checker and Thesaurus](#spell-checker-and-thesaurus)
    - [Software Development](#software-development)
- [The EdSharp Window](#the-edsharp-window)
- [Working With Files](#working-with-files)
- [Working With Text](#working-with-text)
- [Getting Useful Information](#getting-useful-information)
- [Creating an Address Book](#creating-an-address-book)
- [Conclusion](#conclusion)

### Introduction

EdSharp is a full-featured, friendly, powerful, open-source text editor. It uses a standard Windows interface that supports multiple document windows, and it seeks to optimize efficiency for screen reader users by automatically speaking relevant information.

EdSharp works like Notepad, so you can begin using it with the same commands you already know. It adds many more commands that generally involve a modifier key such as Shift, Control, or Alt combined with a letter that begins the name of the command. Most commands provide enhanced speech compared to default screen reader output; for example, EdSharp reads the current line after a command completes. In general, EdSharp does not limit the size of a file or the number of open files, and it includes commands that make it friendly for programmers who want to use it to develop software.

### Installation

The installation program for EdSharp is called EdSharp_Setup.exe. When you run it, it does the following:

- Prompts for a program folder. The default is `C:\Program Files (x86)\EdSharp`.
- Creates a program group on the Windows Start menu with choices to launch EdSharp, read the documentation, or uninstall the program.
- Offers to set or clear an association between EdSharp and files with a particular extension, such as `.txt` or `.ini`. Binary formats such as `.pdf` or `.pptx` may also be associated with EdSharp so they are automatically converted to text when opened from Windows Explorer.
- Presents a checkbox for an optional set of JAWS scripts that fine-tune the EdSharp speech interface. If you install the scripts and later prefer default JAWS behavior, you can disable them from the JAWS "Manage Application Settings" dialog.
- Presents a checkbox that sets Alt+Control+E as a system-wide hot key for the EdSharp shortcut placed on the Windows desktop. If that hot key conflicts with another shortcut, select either shortcut on the desktop, press Alt+Enter for its properties, and change or clear the hot key.
- Presents a checkbox to open the manual in your default web browser after installation.

You can safely install new versions of EdSharp over previous versions; your settings, Favorites and Recent Files lists, and bookmarks are preserved. To update an existing installation, just press F11 (Elevate Version) and EdSharp downloads and installs the current version from the author's site. To check your version at any time, use the About command, Alt+F1; the History of Changes command, Shift+F1, summarizes fixes and improvements over time.

### Quick Start

If you are impatient to get going, here are some things you can do to get up and running.

- Launch EdSharp from the Start menu, the desktop shortcut, or its desktop hot key Alt+Control+E.
- Press F1 to open the full documentation in your default web browser.
- Press Alt+Shift+H for the Hotkey Summary, an alphabetically sorted list of hot keys placed in a new EdSharp window.
- Press Control+F1 for the Key Describer, then press any key to hear which command it runs.
- Press Alt+F10 for the Alternate Menu, an alphabetical list of every available command.
- Explore the regular menus to get a feel for the large number of familiar and new commands available through hot keys.

### Common Hot Keys

Here is a list of hot keys you will use often. Where EdSharp differs from what you may remember from other editors, the EdSharp binding is the one shown.

#### Working With Files on Disk

- Control+N = New file.
- Control+Shift+N = New file from the current clipboard text.
- Control+O = Open file. Converts Microsoft Word (`.doc`/`.docx`), Excel, PowerPoint, PDF, Rich Text Format (`.rtf`), and HTML files automatically to plain text.
- Control+Shift+O = Open Other Format: open a file without conversion (and import `.rtf` with its formatting). Useful for editing HTML source.
- Alt+O = Open Again: reload the current file from disk, discarding unsaved changes.
- Control+S = Save.
- Control+Shift+S = Save As.
- Alt+Shift+S = Save Copy (saves a backup copy under a new name).
- Alt+R = Recent Files list.
- Control+L = Add the current document to the Favorites list.
- Alt+L = Open the Favorites list.
- Alt+Shift+F = File Find: locate a file in the current folder by a text string and open it.
- Control+H = Convert the current document to HTML (HTML Format).
- F5 = Run: open or run the current file with its associated program (for example, an `.htm` file opens in your web browser).
- Control+Tab and Control+Shift+Tab = Cycle forward and backward through open EdSharp windows.
- Control+F4 = Close the current window. Alt+F4 = Exit EdSharp.

#### Navigating in the Current File

- Arrow keys, Page Up, and Page Down = Move through text. Add Shift to select.
- Alt+Down and Alt+Up = Next and prior sentence (EdSharp speaks the sentence).
- Control+Down and Control+Up = Next and prior paragraph (EdSharp speaks the paragraph).
- Control+K = Set a bookmark. Alt+K = Go to a bookmark. Control+Shift+K = Clear the bookmark (the cursor must be on the exact marked character).
- Control+J = Jump to a line number. Alt+J = Jump again. You may add a `+` or `-` before the number to jump that many lines forward or backward from the current line.
- Control+G = Go to a percentage point in the file. Alt+G = Go to that percentage again.

#### Using the Clipboard

- F8 = Start Selection. Move the cursor to one character past the end of the text you want, then press Shift+F8 to Complete Selection.
- Shift plus the arrow keys = Select by character, word, or line in the usual way.
- Control+C = Copy. Control+X = Cut. Control+V = Paste. Each works on the current line if there is no selection.
- Alt+C = Copy and append to the clipboard. Alt+X = Cut and append to the clipboard.
- Alt+Apostrophe = Speak the current clipboard text.

#### Changing Character Case

- Control+U = Upper case the current character or selection.
- Control+Shift+U = Lower case the current character or selection.
- Alt+U = Proper case (capitalize the first letter of each word, lower the rest).
- Alt+Shift+U = Swap case (invert upper and lower case).

#### Searching and Replacing

- Control+F = Find forward. Control+Shift+F = Find backward.
- F3 = Find again forward. Shift+F3 = Find again backward.
- Alt+F3 = Find the chunk at the cursor, or the selected text, forward. Alt+Shift+F3 = the same, backward.
- Control+F3 = Find forward with a regular expression. Control+Shift+F3 = the same, backward.
- Control+R = Replace throughout all or selected text.
- Control+Shift+R = Replace using a regular expression.

#### Working With Sections

EdSharp lets you divide a document into named sections, which is the basis of the address-book exercise later in this tutorial.

- Control+Enter = Insert a Section Break.
- Control+PageDown and Control+PageUp = Move to the next and prior section.
- F6 = Go to Section: jump from a table-of-contents line to its section. Shift+F6 = Go to Contents: jump from a section back to its table-of-contents line.
- Alt+T = Speak the current section's topic (its first line).
- Alt+Shift+T = Text Contents: build a table of contents from the first line of each section.
- Control+F6 = Search for a topic by name. Alt+F6 = Search for that topic again.

#### Spell Checker and Thesaurus

- F7 = Spell check.
- Shift+F7 = Thesaurus.

#### Software Development

- Control+Space = Select Chunk. Good for grabbing a function call or header in one step. Shift+Backspace speaks the current chunk.
- Shift+Enter = Insert a new line indented to the current line's level. Alt+Shift+Enter does the same but opens the line above.
- Tab = Indent the current or selected lines one level. Shift+Tab = Outdent them one level.
- Alt+I = Speak the indentation level of the current line.
- Alt+Shift+I = Indent Mode: toggle the spoken indentation alert on and off.
- Control+Q = Quote the selected text or whole document with a prefix string. You can set the prefix to match the comment string of your programming language to create comments, or use it for quoting email.
- Alt+PageDown and Alt+PageUp = Next and prior part: move to the next or previous function, method, or class, using the current compiler's NavigatePart pattern.
- Control+B and Control+Shift+B = Next and prior block: move by indentation, to the next or previous less-indented line.
- Control+F5 = Compile the current file, speak the output, and jump to the first error. With no compiler configured, a `.cs` file is compiled with the latest available .NET C# compiler automatically. Control+Shift+F5 picks or configures a compiler.

### The EdSharp Window

This section briefly explains the parts of the program you use most often: the title bar, the menu bar, the document area, and the status bar.

The title bar shows the name of the application and, in brackets, the document you are currently editing. The menu bar contains categories of commands you use throughout your work. The document area contains the document you are editing. The status bar reports information about what is happening as you work.

### Working With Files

This section explains how to create, open, and save files.

**Creating a new file.** When EdSharp first opens, it presents a blank document, ready for typing. To create another new file at any time, press Control+N; New is also the first choice on the File menu, Alt+F. To start a file from whatever is currently on the clipboard, press Control+Shift+N: EdSharp opens a new window, drops in the clipboard text, and places the cursor at the start.

**Opening files.** The simplest way to open a file is Control+O, which presents the standard Windows Open dialog showing the current folder. When you open a file this way, EdSharp converts several formats to plain text automatically: Microsoft Word, Excel, PowerPoint, PDF, Rich Text Format, and HTML. To open a file without conversion, use Control+Shift+O (Open Other Format); this is what you want for editing HTML source or importing an `.rtf` file with its formatting intact.

If you are editing a file and want to return to the version on disk, press Alt+O (Open Again). This only helps if you have not already saved your changes, since saving replaces the disk copy.

EdSharp keeps a Recent Files list, reached with Alt+R: pick a file and press Enter to load it. The Favorites list, Alt+L, stores files you return to often, such as a reference manual for a programming project, so you can open them without navigating the Open dialog. To add the file you are editing to Favorites, press Control+L while editing it. EdSharp does not convert files opened from the Recent Files or Favorites lists, so you can re-open a file in its native format; this is most useful for HTML files.

The File Find command, Alt+Shift+F, locates a file in the current folder from a text string: type the string and press Enter, then choose from the resulting list of matching files and press Enter to open it.

**Saving files.** Control+S is the key you will use most often. The first time you save a new file, EdSharp opens the Save As dialog so you can name it; afterward it simply writes the current version over the disk copy. Control+Shift+S (Save As) saves under a different name and switches the document window to the new file. Alt+Shift+S (Save Copy) writes a separate backup copy and leaves you working on the original, which is handy when you want to preserve a snapshot. For more file commands, explore the File menu, Alt+F.

### Working With Text

Besides entering and reviewing text, you will spend a lot of time manipulating it. This section covers the most common and useful operations.

**Selecting text.** Besides the usual Shift-with-navigation method, EdSharp offers two efficient techniques. The first is a two-key selection: place the cursor where the selection should start and press F8; move the cursor one character past the end of the text you want and press Shift+F8 to complete the selection. This is faster than holding Shift, because you do not have to keep a key down or wait for your screen reader to speak as you go. The second technique is Select Chunk, Control+Space, which is excellent for programmers: place the cursor in a function call and press Control+Space to select the entire call at once.

**Changing case.** Use Control+U for upper case, Control+Shift+U for lower case, Alt+U for proper case (the first letter of each word), and Alt+Shift+U to swap (invert) the case of the selection. These commands are also at the bottom of the Misc menu, Alt+M.

**Working with the clipboard.** You already know Control+X to cut and Control+C to copy a selection. EdSharp makes these more efficient: with no selection, Control+X cuts the current line and Control+C copies it, so you can skip selecting. Using Alt+X and Alt+C instead appends the cut or copied text to whatever is already on the clipboard, and these also work on the current line. You will find these and many more commands on the Edit menu, Alt+E.

**Navigating through text.** As you read, you will reach most often for next and prior paragraph (Control+Down and Control+Up), next and prior sentence (Alt+Down and Alt+Up), set and go to bookmark (Control+K and Alt+K), and find forward and backward (Control+F and Control+Shift+F, repeated with F3 and Shift+F3). The remaining navigation commands are on the Navigate menu, Alt+N.

### Getting Useful Information

EdSharp has a set of commands for learning about the document you are working on; most are on the Query menu, Alt+Q.

- Alt+Apostrophe speaks what is currently on the clipboard.
- Alt+A (Address) speaks the line and column of the cursor, followed by its percentage position from the top of the file.
- Alt+Y (Yield) reports how many characters, words, and lines the file, or the selection, contains.
- Alt+Z (Status) tells you whether the document has been modified from the version on disk; press it again to hear the character encoding.

### Creating an Address Book

In this exercise we will use EdSharp's section feature to build a simple address book. Follow along to get a feel for how easily you can manipulate text in the program. We will create a section break and a template for the first entry, add a few fictitious entries, build and sort a table of contents, confirm we can navigate among the entries, convert the book to HTML, and view it in a web browser.

**Create the template entry.**

- Press Control+N to start a new file.
- Press Control+Enter to insert a section break for the template entry.
- Type the following line exactly. Because it begins with an exclamation mark, it will sort to the top when we order the table of contents:

```
!Template
```

- Below that, add a name line in the form we will reuse for every entry:

```
LastName, FirstName
```

- Since this is a simple address book, add only a home address. Put these lines below the name line, with a space after each colon so that you can press End on a line and start typing the value:

```
Street: 
City: 
State: 
Zip: 
Phone: 
Email: 
```

**Make the other entries.**

- Move to the `LastName, FirstName` line of the template.
- Press F8 to start selecting, move to one character past the last line of the template entry, and press Shift+F8 to complete the selection.
- Press Control+C to copy the template to the clipboard.
- Move below the template entry and press Control+Enter to start a new section.
- Press Control+V to paste the template.
- Move to the first line, replace it with a real last name, comma, and first name, for example `Jones, Joe`. Optionally add a middle initial by pressing End, Space, and a capital letter.
- Fill in the remaining lines. As long as you do not change the clipboard, you can keep pasting the template (Control+V) under new section breaks to add more entries. Create a few until the process feels comfortable.

**Build the table of contents.** Press Alt+Shift+T (Text Contents). EdSharp adds a "Contents" line near the top of the file followed by the first line of each section, which here is the name line of each entry. A blank line in the table of contents means a section break is missing or misplaced.

**Check the sections.** Use Control+PageDown and Control+PageUp to confirm the cursor lands on the first line of each entry. If you land on a blank line, delete any text between the previous entry's last line and the new entry's first line, place the cursor on the first character of the new entry, and press Control+Enter to recreate the section break. When the sections are correct, make sure the table of contents has a line for each one.

**Sort the table of contents.** Select all of the table-of-contents lines and press Alt+Shift+O (Order Items) to sort them alphabetically. The `!Template` line stays at the top because of its leading exclamation mark. Now you can press F6 on a contents line to jump to that entry, and Shift+F6 to jump back to the contents.

**Convert to HTML and view it.** Press Control+H (HTML Format) to convert the file to HTML. Save it with Control+S, then press F5 (Run) to open it in your default web browser. The page should show same-page links to each address-book entry. If it does not, check the section breaks and the table of contents, and compare the structure against the examples in the EdSharp documentation (F1).

### Conclusion

This tutorial has covered the commands you will use most often in day-to-day work with text, and it has shown how the section feature lets you build and navigate a structured document such as an address book and convert it to HTML. For the complete set of commands and features, including programming, math, word processing, and scripting, press F1 in EdSharp to open the full User Guide. EdSharp's author welcomes feedback and contributions toward its continued improvement.

<!-- walkthrough: written by makeTutorials.py, do not edit between the markers -->

## 00 - Overview and Table of Contents

A simulated walk through EdSharp: a person working, and a screen reader answering.

**Before you start:** Nothing to prepare. This one is a few minutes on what EdSharp is and where the rest of the walks go.

### Step 1

EdSharp is a text editor built for people who listen to their screen rather than look at it. It opens plain text, Markdown, HTML and program source, and it converts between about forty formats without leaving the keyboard.

### Step 2

Everything here is a key. There are no toolbars to hunt for and no panes to get lost in. Where a key has a letter in it, the letter is the first letter of a word in the command's name, so the key is something you work out rather than something you memorise.

### Step 3: Alt+Control+E

Let us start the program and hear what it says. The editor opens with one empty document.

Screen reader:

- EdSharp
- Untitled, edit, multiline, blank

### Step 4: F1

Press F1 for Help. F1 is Help everywhere in Windows, and in EdSharp it opens the user guide.

Screen reader:

- EdSharp User Guide, document

### Step 5: Insert+T

Come back to the editor with Alt plus Tab, then let us hear where we are. Insert plus T, Tango, says the window title.

Screen reader:

- EdSharp

### Step 6: Insert+UpArrow

If you miss what was said, Insert plus Up Arrow says the current line again. That key is worth learning first, because it turns every other key in these walks into something you can replay.

Screen reader:

- Untitled, edit, multiline, blank

### Step 7: Alt+F10

Now the command space. Alt plus F10 opens the Alternate Menu, which lists every command EdSharp has, with its key beside it.

Screen reader:

- Alternate Menu
- Command, list box, 1 of 222

### Step 8: c o n v e r t

Type the word convert and listen.

Screen reader:

- Convert File Format, Control plus Shift plus V, 1 of 9

### Step 9: Escape

Escape closes the menu without running anything.

Screen reader:

- Untitled, edit, multiline, blank

### Step 10: F11

One more key before the tour of the other walks. F11 is Elevate Version. Elevate sounds like eleven, which is how the key was chosen.

Screen reader:

- Elevate Version, dialog
- EdSharp 5.0 is the newest version. Yes button.

### Step 11: Escape

Nothing was changed.

Screen reader:

- Untitled, edit, multiline, blank

### Step 12

That is the shape of it. The other walks each take one thing and go through it slowly.

The walks below are the planned set. Each becomes its own Tutorial_NN file.

### Step 13

Walk one is opening, editing and saving, including the encoding of a file and what to do when it opens as nonsense.

### Step 14

Walk two is moving around a long document by headings, by blocks that mean something in the language you are in, and by bookmarks.

### Step 15

Walk three is converting between formats: Markdown to a web page, a web page to text, a PDF to something you can read.

### Step 16

Walk four is writing code: checking the syntax without running it, stepping definition to definition, and the snippets that fill in the parts that change.

### Step 17

Walk five is the local AI: chatting about the document in front of you, and translating it, with no web service involved.

### Step 18

Pick the one that matches what you want to do today. Each is about three minutes, and each ends with something to try.

**Something to try:** Open a file you already have and press F1. Then come back and pick the walk that matches what you want to do today.

<!-- walkthrough ends -->
