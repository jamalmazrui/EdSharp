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

What EdSharp is, in a paragraph; the two reader keys every walk assumes; then the table of contents, one line per walk.

**Before you start:** Nothing is needed; this walk is listened to.

### Step 1

EdSharp is a text and code editor for people who listen to the screen rather than look at it. It opens plain text, Markdown, HTML and program source; it opens Word documents, PDF files, slide decks and spreadsheets as text you can read and search; it compiles and runs programs and lands the cursor on the error; it checks spelling; and it talks to an AI model on your own computer. Everything is a key, and every key is named for a word in its command.

### Step 2: Insert+UpArrow

Two keys before anything else, both the reader's own. If a line goes by too fast, Insert plus Up Arrow says it again.

Screen reader:

- (the last line, read a second time)

### Step 3: Insert+Tab

And if you lose your place, Insert plus Tab says where you are: the control, its state, its position, and any hint it carries. The walks are heard with those hints off, the way most people work.

Screen reader:

- (the current control, with its state and position)

### Step 4

Now the table of contents. I say the number and the title; the reader says what the walk covers.

### Step 5

One, Install and Launch.

Screen reader:

- The download, the installer's last page and its boxes, and EdSharp opening with Alt plus Control plus E.

### Step 6

Two, User Interface Concepts.

Screen reader:

- One window, one document at a time with others behind it, the edit box you hear, dialogs that all work one way, and where help is.

### Step 7

Three, Key Patterns.

Screen reader:

- The rules every key follows, so a key can be guessed before it is learned, and the keys that explain the keys.

### Step 8

Four, Open, Edit and Save a File.

Screen reader:

- A file opened, a line changed, the document saved, and the two ways to hear where you are.

### Step 9

Five, Find Your Way in a Long Document.

Screen reader:

- Finding a word, jumping to a line, and bookmarks that bring you back.

### Step 10

Six, Convert Between Formats.

Screen reader:

- A Word document opened as text; Markdown written out as a web page.

### Step 11

Seven, Write Code.

Screen reader:

- A program compiled from the editor, the cursor landing on the error, and blocks of code moved through.

### Step 12

Eight, Ask the AI on Your Own Computer.

Screen reader:

- A question answered and a document translated, with nothing sent off the machine.

### Step 13

Nine, Glossary.

Screen reader:

- The words EdSharp uses, in alphabetical order, one line each.

### Step 14

Ten, Conclusion.

Screen reader:

- Four sentences to carry away, and where to begin.

### Step 15

Eleven, More Information.

Screen reader:

- The guide and history from inside EdSharp, the documents, the GitHub page, updates, and the other Homer Tools.

### Step 16

Twelve walks, each under five minutes, about an hour together. They are a course, not a reference: each one assumes those before it.

**Something to try:** Listen to the walks in order; each one assumes the ones before it.

## 01 - Install and Launch

The download, the installer's pages, the last page with its boxes and its summary, and EdSharp opening by itself.

**Before you start:** The installer is downloaded, and your reader is running.

### Step 1: Enter

The installer is downloaded; Enter opens it. Windows may first ask whether to run a file from the Internet; Alt plus R, Run, answers it.

Screen reader:

- EdSharp Setup dialog

### Step 2: Enter

Each page is a dialog like any other: Tab through it, Enter for Next. Accept the defaults; the one page that asks a decision is the last.

Screen reader:

- Next button

### Step 3

The last page offers optional pieces as boxes, each saying what it does and how much space it needs. Python and the document tools are ticked: about a hundred and fifty megabytes, for rich PDF conversion and the thesaurus. Git, Node and Ollama with its AI models are not ticked; each backs a real feature, and each serves some people and not others. Spacebar changes a box.

### Step 4: Enter

Enter on Finish. Whatever was ticked installs now, and one summary then reports each item by name. The summary is saved in the logs folder, and summarizeSetup, in the program folder, shows it again.

Screen reader:

- EdSharp Setup, Install Python: installed. Install document tools: installed.

### Step 5: Alt+Control+E

EdSharp opens by itself. From then on, Alt plus Control plus E opens it from anywhere in Windows, or brings it forward if it is already open.

Screen reader:

- EdSharp
- Untitled, edit, multiline, blank

### Step 6

What each box is for. Python runs the conversion tools and the spell checker's helpers; the document tools -- pandoc and its companions -- turn Word, PDF and web pages into text and back. Together they are what makes Control plus Shift plus O and Alt plus Shift plus E rich rather than basic.

### Step 7

Git is for people who keep code in repositories; Node for web developers who run JavaScript tools; Ollama for the AI features, with a model that is a few gigabytes. Each box says its size, so the decision is yours and informed.

### Step 8

A box ticked by mistake is no disaster: the same installer, run again later, offers the same page, and an item already present says so and does nothing. Nothing is installed twice.

### Step 9

The screen reader scripts are offered too. Install when they are not there, Update when a newer set is available, Reinstall when they are current -- the word is the state, so you need not ask.

### Step 10

Where things went. EdSharp itself is in Program Files; your settings, snippets and logs are under your local application data, in a folder named EdSharp. An update never touches them.

### Step 11: F11

F11, Elevate Version, is the installer's other half from now on: it asks the web for a newer EdSharp and offers to fetch and run it. Elevate sounds like eleven.

Screen reader:

- Elevate Version dialog
- EdSharp 5.0 is the newest version. Yes button

### Step 12: Escape

Escape. Alt plus F1 says the version at any time; Control plus F12 copies the session log's path, ready to attach to a report of a problem.

Screen reader:

- Untitled, edit, multiline, blank

### Step 13

Before the installer, the download: the ReadMe on the project page has one link, EdSharp underscore setup dot exe, and F11 inside an installed EdSharp fetches the same file.

### Step 14

A planned misstep. Run the installer while EdSharp is open and it asks you to close it first, by name; close it, press Retry, and the installer carries on. Nothing is half-installed.

### Step 15

The first run asks nothing. EdSharp opens on an empty document, and the reader says the title, then the edit box, then blank; typing starts at once.

### Step 16

What the screen reader scripts give. With them installed, the reader knows EdSharp's edit box and dialogs by name, and reads the status line's messages as they arrive without your asking.

### Step 17

What to try first. Type a sentence, Control plus S, give it a name, Alt plus A to hear where you are, F1 for the guide. Five minutes of that and walk four is half known.

### Step 18

Updating later is F11 inside EdSharp: it says the installed version and the newest one, and offers to fetch and run the installer, whose last page shows the same boxes, each saying its state.

### Step 19

What this walk taught. I say the key; the reader says what it does.

### Step 20

Alt plus Control plus E.

Screen reader:

- Open EdSharp

**Something to try:** Install EdSharp and open it with Alt plus Control plus E.

## 02 - User Interface Concepts

What EdSharp is made of: one window with documents behind it, the edit box you hear, dialogs that all work one way, the two keys that say where you are, messages, and where help is.

**Before you start:** EdSharp is open. Nothing needs pressing in this walk; it is listened to.

### Step 1

EdSharp is one window. The document you are editing fills it; others you have opened wait behind it. Control plus Tab moves to the next; F4 lists them all by name.

### Step 2

What you hear is the edit box: the window title, then Untitled or the file's name, then edit, multiline, and the line the cursor is on. It is a plain Windows edit control, so every reader reads it the same way.

### Step 3

Every dialog is built the same way: a label and its control, Tab between them, Alt plus the underlined letter to jump to one, Control plus Enter for OK from anywhere, Escape to cancel. Open and Save are the Windows dialogs you already know.

### Step 4

Two keys say where you are. Alt plus A, Say Address, says the line, the column and how far down the document you are. Alt plus Z, Say Status, says whether the document has changed since it was saved; pressed again, its character encoding.

### Step 5

Messages from EdSharp are spoken as they happen, and shown on the status line at the bottom: Converting, Line, Bookmark at percent 40. Insert plus Page Down, the reader's own key, reads that line.

### Step 6

Help is in four places, and they are the same in every Homer program. F1 opens the guide, the whole program in one document. Shift plus F1 opens the history of changes. Alt plus F1 says the version and offers the newer one if there is one. Control plus Shift plus F1 opens the quick starts by role: Python developer, screen reader script writer, translator, slide presenter and more.

### Step 7

The Help menu, F10 then H, Hotel, holds the same, and Play Tutorials, which plays these walks. And the menus themselves are help: arrow through any menu and the reader says each command with its key and its letter.

### Step 8

Three kinds of window besides the document. A result window -- the output of a compile, the answer from the AI, a hotkey summary -- is a document like any other, read-only until you decide otherwise. A preview window shows Markdown formatted. A dialog asks and closes.

### Step 9

Text boxes that ask for something -- a search, a file, a line number -- remember your last answers: Up and Down Arrows in the box bring them back, and the newest is offered first.

### Step 10

Selection is spoken as you make it: Shift with the arrows selects, and Shift plus Space, Say Selected, reads what is selected; twice, it spells. Control plus Space selects the chunk at the cursor -- a word with no spaces in it -- and again selects the next.

### Step 11

EdSharp speaks its own confirmations so you need not ask: Converting when a conversion starts, the count of lines after a sort or a wrap, Bookmark at percent 40 when a bookmark is set. If a message is missed, Insert plus Page Down reads the status line, and Alt plus Z says whether the document is modified.

### Step 12

Formats are a matter of the file's extension. A dot m d file is Markdown and gets the preview; a dot p y file is Python and gets its compiler; a dot h t m file is a web page and opens in the browser with F5. Save As with a new extension changes what EdSharp offers.

### Step 13

Settings live in one dialog, Alt plus Shift plus C, Configuration Options: the spell checker, the font, what is said and what is not. Each is a labelled control, so a setting is found by its name.

### Step 14

What this walk taught. I say the key; the reader says what it does.

### Step 15

F1.

Screen reader:

- Documentation

### Step 16

Alt plus A.

Screen reader:

- Say Address

### Step 17

Alt plus Z.

Screen reader:

- Say Status

**Something to try:** Open EdSharp and name each thing as you reach it: the title, the edit box, the status line.

## 03 - Key Patterns

The rules every EdSharp key follows, so a key can be guessed before it is learned: the word gives the letter, Control does, Alt says, Shift widens, Alt Shift is a command with no control, the function keys follow Windows -- and the keys that explain the keys.

**Before you start:** EdSharp is open on an empty document.

### Step 1

Every key in EdSharp is named for a word in its command: Control plus O is Open, Control plus S is Save, Control plus F is Find, Control plus K is bookmark, F7 is spell check as in Word. A key never comes from the middle of a word.

### Step 2

Control plus a letter does something. Alt plus a letter says something and changes nothing: Alt plus A the address, Alt plus Z the status, Alt plus P the path, Alt plus Y the yield -- the count of characters, words and lines.

### Step 3

Adding Shift widens or reverses a key. Control plus O opens a file; Control plus Shift plus O opens one in another format, converting it to text. Control plus F finds forward; Control plus Shift plus F finds backward. F3 finds the next; Shift plus F3 the previous.

### Step 4

Alt plus Shift plus a letter is a command with no control of its own: Alt plus Shift plus E exports the document as another format, Alt plus Shift plus H shows the hotkey summary.

### Step 5

The function keys follow Windows and Office: F1 help, F3 find again, F4 the open documents, F5 run, F7 spelling, F11 the newer version, F12 the AI. Control plus F5 compiles; Control plus Shift plus F5 chooses the compiler.

### Step 6: Alt+F10

The keys that explain the keys. Control plus F1 is the Key Describer: on, every key says what it does instead of doing it, the safe way to explore the keyboard. Alt plus F10 is the Alternate Menu: every command in one list you type into, and the list tells you the key.

Screen reader:

- Alternate Menu
- Command, list box, 1 of 222

### Step 7: convert

Type the word convert.

Screen reader:

- Convert File Format, Control plus Shift plus V, 1 of 9

### Step 8: Escape

Escape closes the menu without running anything. Hotkeys, in the Help menu, lists every key three ways; and your reader's own Insert plus Tab says where you are.

Screen reader:

- Untitled, edit, multiline, blank

### Step 9

The rules let a key be guessed. I say a command; the reader says the key the rules give it.

### Step 10

Save As.

Screen reader:

- Control plus Shift plus S

### Step 11

Replace.

Screen reader:

- Control plus R

### Step 12

Go to a line by number -- Jump to Line.

Screen reader:

- Control plus J

### Step 13

Say the path of this file.

Screen reader:

- Alt plus P

### Step 14

Say the time.

Screen reader:

- Alt plus Semicolon

### Step 15

Insert the time.

Screen reader:

- Alt plus Shift plus Semicolon

### Step 16

Copy the whole document.

Screen reader:

- Control plus F8

### Step 17

Keys that are the reader's are not EdSharp's: anything with Insert in it belongs to the screen reader, and EdSharp never uses the Insert key, so the two never collide.

### Step 18

One more family. Alt plus a letter says; Alt plus Shift plus the same letter often acts on the same thing: Alt plus Y says the yield, Alt plus Shift plus Y renders the text in another encoding; Alt plus V invokes a snippet, Alt plus Shift plus V views one.

### Step 19

What this walk taught. I say the key; the reader says what it does.

### Step 20

Alt plus F10.

Screen reader:

- Alternate Menu

### Step 21

Control plus F1.

Screen reader:

- Key Describer

**Something to try:** Guess the key for Replace, Save As and Thesaurus before you look them up, then check with Control plus F1.

## 04 - Open, Edit and Save a File

One want: a letter to finish. Opening it, moving through it, changing it, hearing where you are and whether it is saved, checking the spelling, and saving. This walk assumes walks one to three.

**Before you start:** EdSharp is open, and letter.txt is in your Documents folder.

### Step 1: Control+O

The want: a letter to finish. Control plus O, Open, is the Windows open dialog; type the name, or arrow the list.

Screen reader:

- Open dialog, File name: edit combo

### Step 2: Enter

letter dot t x t, and Enter.

Screen reader:

- letter dot t x t, EdSharp
- Dear Ms. Alvarez,

### Step 3: DownArrow

Arrow down through it as in any editor. Control plus Down Arrow is the next paragraph; Control plus Right Arrow the next word, read as you land on it.

Screen reader:

- Thank you for your letter of the fourth.

### Step 4: Alt+A

Type at the end, Shift plus Enter for a new line keeping the indentation, Enter for a plain one. Then Alt plus A, Say Address, says where you are.

Screen reader:

- Line 2, column 42, 30 percent

### Step 5: Alt+Z

Alt plus Z, Say Status: has the document changed since it was saved?

Screen reader:

- Modified

### Step 6: Control+S

Control plus S saves. The title loses its mark and the status becomes unmodified; nothing more is said, because nothing more happened.

Screen reader:

- letter dot t x t, EdSharp

### Step 7: F7

F7 checks spelling, all or selected text. Each misspelling comes as a dialog: the word, its context, the suggestions in a list; Enter takes the one you are on, Escape leaves the word alone.

Screen reader:

- Spell Check dialog, recieve, list box, receive, 1 of 3

### Step 8: Enter

Enter takes receive, and the next misspelling follows, until the end.

Screen reader:

- Spell check complete

### Step 9: Control+S

Shift plus F7, Thesaurus, lists synonyms for the word at the cursor, the same way: a list, Enter to replace. Control plus S again, and the letter is done.

Screen reader:

- letter dot t x t, EdSharp

### Step 10: Control+Z

A planned misstep. I meant to type in the letter and typed over a word I had selected. Control plus Z, Undo, as everywhere in Windows.

Screen reader:

- Thank you for your letter of the fourth.

### Step 11

A new file starts with Control plus N, New; the first Save asks for a name, and the extension you give it decides what EdSharp will offer -- dot t x t for plain text, dot m d for Markdown.

### Step 12

Control plus Shift plus S, Save As, saves under a new name, and the window takes it; Alt plus Shift plus S, Save Copy, writes a copy and leaves you in the original.

### Step 13: F4

Control plus F4 closes the document; if it has changed, EdSharp asks once. Control plus Tab moves to the next open document; F4 lists them all, so a dozen files are a list away.

Screen reader:

- Current Windows dialog, list box, letter dot t x t, 1 of 2

### Step 14: Alt+P

Escape. Alt plus P, Say Path, says where this file is on disk, for when the title is not enough.

Screen reader:

- C colon, Users, Jamal, Documents, letter dot t x t

### Step 15

What this walk taught. I say the key; the reader says what it does.

### Step 16

Control plus O.

Screen reader:

- Open

### Step 17

Control plus S.

Screen reader:

- Save

### Step 18

F7.

Screen reader:

- Spell Check

### Step 19

Alt plus A.

Screen reader:

- Say Address

**Something to try:** Open a file of your own, change one line, check its spelling, and save it.

## 05 - Find Your Way in a Long Document

One want: a place in a long report, found now and found again tomorrow. Finding a word, jumping to a line, setting a bookmark and returning to it, and hearing how big the document is. This walk assumes walk four.

**Before you start:** A long document is open in EdSharp.

### Step 1: Control+F

The want: a place in a long report, found again tomorrow. Control plus F, Forward Find, asks for the text and remembers your last answers.

Screen reader:

- Forward Find dialog, Find: edit combo

### Step 2: Enter

budget, and Enter. The cursor lands on the first match, and the line is read.

Screen reader:

- The budget for the second quarter is attached.

### Step 3: F3

F3 finds the next; Shift plus F3 the previous. A find that reaches the end says so and stops.

Screen reader:

- Budget notes follow in the appendix.

### Step 4: Control+J

Control plus J, Jump to Line, takes a number: a line, or a line and a column.

Screen reader:

- Jump to Line dialog, Line: edit

### Step 5: Enter

120, and Enter.

Screen reader:

- Section four. Staffing.

### Step 6: Control+K

Control plus K, Set Bookmark, marks the place; EdSharp says how far down the document it is.

Screen reader:

- Bookmark at percent 40

### Step 7: Alt+K

Move away -- Control plus Home to the top -- then Alt plus K, Go to Bookmark, brings you back. Bookmarks are kept with the file, so tomorrow's Alt plus K lands here too.

Screen reader:

- Section four. Staffing.

### Step 8: Alt+Y

Alt plus Y, Say Yield, says the size of what you are in: characters, words and lines, for all the text or the selection.

Screen reader:

- 14,205 characters, 2,310 words, 188 lines

### Step 9: Control+R

Control plus R, Replace, is Find's sibling: the text to find, the text to put in its place, in all the document or in the selection. Each replacement is counted and the count is said.

Screen reader:

- Replace dialog, Find: edit combo

### Step 10: Escape

Escape. For a pattern rather than a word -- every number, every line that begins with a dash -- Control plus F3 finds with a regular expression, and Control plus Shift plus R replaces with one. The expressions are the dot NET kind, and the guide has a page on them.

Screen reader:

- Section four. Staffing.

### Step 11

The document's shape is a set of keys too. Control plus Down Arrow is the next paragraph; Alt plus Right Arrow the next chunk, read as you land on it; Control plus Right Arrow the next word.

### Step 12: Alt+A

Alt plus A, Say Address, is the anchor for all of this: line, column and percent, at any moment, so you know where a find landed you before you start to type.

Screen reader:

- Line 120, column 1, 64 percent

### Step 13: Alt+Shift+F

Alt plus Shift plus F, File Find, searches not this document but a folder of files for a string, and opens the one you pick from the list -- for the day the place you want is in some other file.

Screen reader:

- File Find dialog, Find: edit combo

### Step 14: Escape

Escape.

Screen reader:

- Section four. Staffing.

### Step 15

What this walk taught. I say the key; the reader says what it does.

### Step 16

Control plus F.

Screen reader:

- Forward Find

### Step 17

Control plus J.

Screen reader:

- Jump to Line

### Step 18

Control plus K.

Screen reader:

- Set Bookmark

### Step 19

Alt plus K.

Screen reader:

- Go to Bookmark

**Something to try:** Find a word in a document of your own, bookmark the place, go to the top, and come back.

## 06 - Convert Between Formats

Two wants: a Word document read as text, and Markdown written out as a web page. Opening in another format, exporting to one, and the Markdown preview. This walk assumes walks four and five.

**Before you start:** EdSharp is open, and report.docx is in your Documents folder.

### Step 1: Control+Shift+O

The want: a Word document someone sent, read as text. Control plus Shift plus O, Open Other Format -- Shift widens Open -- converts it on the way in.

Screen reader:

- Open Other Format dialog, File name: edit combo

### Step 2: Enter

report dot d o c x, and Enter. EdSharp says what it is doing.

Screen reader:

- Converting
- report dot d o c x, EdSharp
- Quarterly Report

### Step 3: Alt+Shift+E

The other direction: writing out as something else. Alt plus Shift plus E, Export Format, offers the formats this document can become.

Screen reader:

- Export Format dialog, Format: list box, Web page, 1 of 8

### Step 4: Enter

Web page, and Enter: the Markdown you are writing becomes an HTML file beside it, and EdSharp says the file.

Screen reader:

- Converted 1 file

### Step 5: Control+F9

Control plus F9, Preview Markdown, shows the same thing as a formatted page in a window of EdSharp's own, for checking headings and links without leaving the editor.

Screen reader:

- Preview, document

### Step 6: Escape

Escape returns. PDF files, slide decks and spreadsheets open the same way, by Control plus Shift plus O; the document tools from the installer's last page are what make the richer ones possible.

Screen reader:

- report dot d o c x, EdSharp

### Step 7: Control+Shift+V

Control plus Shift plus V, Convert File Format, converts a file on disk without opening it: pick the file, pick the format, and the new file lands beside the old. Several files at once make a batch, and EdSharp says how many were converted.

Screen reader:

- Convert File Format dialog, File name: edit combo

### Step 8: Escape

Escape. The formats on offer depend on what was installed: the document tools from the installer's last page add Word, PDF, EPUB and slide decks; without them, text, Markdown, HTML and RTF are always there.

Screen reader:

- report dot d o c x, EdSharp

### Step 9

A planned misstep. A PDF that is only a picture of a page -- a scan -- opens as empty or as nonsense, because there is no text in it to convert. The guide's section on PDF says what to do: optical character recognition, which the document tools can run on it.

### Step 10

Encoding is the other kind of conversion. A file that opens as symbols where accents should be is in an encoding EdSharp guessed wrong; Alt plus Shift plus Y, Yield Encoding, renders the text in the one you choose, and Alt plus Z twice says the encoding in use.

### Step 11

Markdown to plain text, HTML to Markdown, a table in a document to lines you can read -- each is an export or an open-other-format away, and the preview, Control plus F9, is the check before you send.

### Step 12

Where converted files go. Open Other Format reads the original and makes a text copy in your temporary folder; Save puts it where you say. Export writes beside the original, with the new extension; Convert File Format does the same for files on disk.

### Step 13

Tables come through as lines: each row is a line, each cell separated by a tab or a bar, which a reader reads better than a grid. Headings keep their level as Markdown's pound signs, so Control plus B and the heading keys still move by them.

### Step 14

A planned misstep. Export to PDF with the document tools not installed, and EdSharp says so and names the box on the installer's last page that provides them; nothing hangs and nothing half-converts.

Screen reader:

- PDF export needs the document tools. Run the installer and tick Document tools.

### Step 15

Images in a document are listed by their alternative text, when the author gave any; a document with none says image and its file name, which is itself a finding when you are checking someone's accessibility.

### Step 16

Round trips are safe: a Markdown file exported to Word and opened again as text returns as the same Markdown, headings and lists intact; what is lost is only what Markdown cannot say, such as a font.

### Step 17

What this walk taught. I say the key; the reader says what it does.

### Step 18

Control plus Shift plus O.

Screen reader:

- Open Other Format

### Step 19

Alt plus Shift plus E.

Screen reader:

- Export Format

### Step 20

Control plus F9.

Screen reader:

- Preview Markdown

**Something to try:** Open a Word document or a PDF as text, and export a Markdown file as a web page.

## 07 - Write Code

One want: a program that will not run, fixed. Choosing the compiler once, compiling with the cursor landing on the error, moving by blocks of code, hearing indentation, and snippets. This walk assumes walks four and five.

**Before you start:** EdSharp is open on a short Python file with one mistake, and Python was ticked when EdSharp was installed.

### Step 1: Control+Shift+F5

The want: a Python program that will not run, fixed. With hello dot p y open, Control plus Shift plus F5, Choose Compiler, picks the language once; EdSharp remembers it by extension.

Screen reader:

- Choose Compiler dialog, Compiler: list box, Python, 7 of 20

### Step 2: Control+F5

Enter. Then Control plus F5, Compile: the program is run or checked, the output is spoken, and the cursor lands on the first error.

Screen reader:

- Compiler Python
- line 12, SyntaxError: expected colon
- for item in items

### Step 3: Control+B

Code is moved through by its shape. Control plus B, Next Block, goes to the next block of code with the same or less indentation; Control plus Shift plus B the previous.

Screen reader:

- def total(items):

### Step 4: Alt+I

Alt plus I, Say Indentation, says the level of the current line; Alt plus B, Say Block, reads the rest of the block, so you hear a loop or a function whole.

Screen reader:

- Indent 1

### Step 5: Control+F5

Fix the line, Control plus S, Control plus F5 again. A clean run speaks its output and lands nowhere.

Screen reader:

- Compiler Python
- Total: 42

### Step 6: Alt+V

Snippets fill in the parts that change. Alt plus V, Invoke Snippet, picks a saved piece of code to paste or run; Alt plus S saves the selection as a new one.

Screen reader:

- Invoke Snippet dialog, Snippet: list box, python-main, 1 of 14

### Step 7: Escape

F5, Run, executes the file by its extension, without the compile step, once it is right.

Screen reader:

- hello dot p y, EdSharp

### Step 8: Alt+0

Alt plus 0, Say Compiler, says which compiler is set for this file and the folder it runs in, so a wrong one is heard before it runs.

Screen reader:

- Compiler Python, folder C colon, Users, Jamal, Code

### Step 9

A planned misstep. Control plus F5 with no compiler chosen for a new extension says so and names the key to fix it.

Screen reader:

- No compiler configured. Press Control plus Shift plus F5 to pick one.

### Step 10

Alt plus Shift plus right bracket, Say Braces, counts the braces on either side of the cursor -- the quickest answer to a parser's complaint about an unmatched one.

### Step 11

Indentation is spoken as structure. Alt plus Shift plus I, Indent Mode, makes Enter keep the current indentation and announces changes of level as you arrow; Tab and Shift plus Tab indent and outdent a line or a selection by one unit.

### Step 12

Shift plus F5, Run at Cursor, runs the web address or email address at the cursor -- a link in a comment, a URL in a string -- in the program Windows has for it.

### Step 13

Alt plus Shift plus F9, Run Code Blocks, runs the SQL and JScript blocks in a Markdown document and writes each block's result beneath it, which turns a notes file into a small notebook.

### Step 14

What this walk taught. I say the key; the reader says what it does.

### Step 15

Control plus F5.

Screen reader:

- Compile

### Step 16

Control plus Shift plus F5.

Screen reader:

- Choose Compiler

### Step 17

Control plus B.

Screen reader:

- Next Block

### Step 18

Alt plus V.

Screen reader:

- Invoke Snippet

**Something to try:** Compile a program of your own, go to its first error, fix it, and run it.

## 08 - Ask the AI on Your Own Computer

One want: a question about a document, answered with nothing sent off the machine. Chatting about the document, chatting with no document, and translating it, all with the model Ollama runs on your own computer. This walk assumes walks four to six.

**Before you start:** EdSharp is open on a document, and Ollama with a model was ticked when EdSharp was installed, or installed since.

### Step 1

The want: a question about the document, answered without sending it anywhere. The model runs on your own computer, through Ollama, which the installer offered; nothing leaves the machine.

### Step 2: Shift+F12

Shift plus F12, Chat about Document: the question goes with the document, or with the selection when text is selected, and the answer opens in a new window.

Screen reader:

- Chat about Document dialog, Question: edit

### Step 3: Enter

What are the three main points? And Enter. EdSharp names the model it is asking.

Screen reader:

- Asking llama3.2
- Answer - EdSharp
- The report makes three points. First,

### Step 4

F12 alone, Chat with AI, asks a question with no document attached. The answer opens the same way.

### Step 5: Alt+Shift+F7

Alt plus Shift plus F7, Translate Language, translates the selection or the whole document between two languages, with the same model. EdSharp names the model and the languages.

Screen reader:

- Translate Language dialog, From: combo, English

### Step 6: Control+Enter

Spanish as the target, Control plus Enter, and the translation opens in a new window, the original untouched.

Screen reader:

- Translating with llama3.2
- Translation - EdSharp

### Step 7

All of it waits on the machine, not on a service: the first answer from a model takes a few seconds while it loads, and the next ones are quick.

### Step 8

Which model answers is a setting: Alt plus Shift plus C, Configuration Options, names the model, and any model Ollama has pulled can be chosen. A small model answers in seconds; a large one answers better and slower.

### Step 9

What makes a good question here is what makes one anywhere: say what you want back -- three bullet points, a one-paragraph summary, the dates mentioned -- and the answer takes that shape. A vague question gets a vague answer, locally as much as online.

### Step 10

The answer window is a document. Alt plus C, Copy Append, adds a passage to the clipboard without losing what is there; Control plus S saves the whole answer beside the document it is about.

### Step 11

A planned misstep. Shift plus F12 with Ollama not running, or no model pulled, says so plainly and names what to do; nothing hangs. Ollama starts with Windows once installed, so this is rare after the first day.

### Step 12

Summarizing is a question: select nothing, Shift plus F12, and ask for a summary in five sentences. Translation is a question too, but Alt plus Shift plus F7 asks it with the languages as a dialog, so the two languages are remembered next time.

### Step 13

Nothing in this walk needs the Internet. The model, the question and the document stay on the machine, which is what makes it fit for a letter you would not paste into a web page.

### Step 14

What this walk taught. I say the key; the reader says what it does.

### Step 15

Shift plus F12.

Screen reader:

- Chat about Document

### Step 16

F12.

Screen reader:

- Chat with AI

### Step 17

Alt plus Shift plus F7.

Screen reader:

- Translate Language

**Something to try:** Ask a question about a document of your own, then translate a paragraph of it.

## 09 - Glossary

The words EdSharp uses, in alphabetical order. I say the term; the reader says what it means. Each is one line, for looking up or for listening straight through.

**Before you start:** Nothing is needed.

### Step 1

address.

Screen reader:

- Where the cursor is: line, column and percent of the document. Alt plus A.

### Step 2

Alternate Menu.

Screen reader:

- Every command in one list you type into, with its key beside it. Alt plus F10.

### Step 3

block.

Screen reader:

- A stretch of code at one indentation. Control plus B goes to the next; Alt plus B reads the rest of this one.

### Step 4

bookmark.

Screen reader:

- A place in a file, kept with the file. Control plus K sets one; Alt plus K returns to it.

### Step 5

chunk.

Screen reader:

- A run of characters with no space in it. Alt plus Right Arrow goes to the next; Control plus Space selects it.

### Step 6

compiler.

Screen reader:

- The program EdSharp hands your file to with Control plus F5, chosen once by extension with Control plus Shift plus F5.

### Step 7

configuration options.

Screen reader:

- The one dialog of settings -- spell checker, font, speech, the AI model. Alt plus Shift plus C.

### Step 8

convert file format.

Screen reader:

- Converting a file on disk without opening it, one or many at a time. Control plus Shift plus V.

### Step 9

current windows.

Screen reader:

- The list of open documents, to pick one. F4; Shift plus F4 says their titles.

### Step 10

document.

Screen reader:

- One file open in EdSharp, in a window of its own; Control plus Tab moves to the next, F4 lists them.

### Step 11

Elevate Version.

Screen reader:

- Checking for a newer EdSharp and installing it. F11, because elevate sounds like eleven.

### Step 12

encoding.

Screen reader:

- How a file's characters are stored. Alt plus Z twice says it; Alt plus Shift plus Y renders the text in another one.

### Step 13

export.

Screen reader:

- Writing the document out as another format -- a web page, a Word document, plain text. Alt plus Shift plus E.

### Step 14

file find.

Screen reader:

- Searching a folder of files for a string and opening the one you pick. Alt plus Shift plus F.

### Step 15

hotkey summary.

Screen reader:

- Every command with its key and description, in a new window. Alt plus Shift plus H.

### Step 16

Key Describer.

Screen reader:

- A mode in which every key says what it does instead of doing it. Control plus F1 turns it on and off.

### Step 17

Ollama.

Screen reader:

- The program that runs AI models on your own computer; EdSharp's F12, Shift plus F12 and Alt plus Shift plus F7 talk to it.

### Step 18

open other format.

Screen reader:

- Opening a Word document, PDF, slide deck or spreadsheet converted to text. Control plus Shift plus O.

### Step 19

preview.

Screen reader:

- Markdown shown as a formatted page in a window of EdSharp's own. Control plus F9.

### Step 20

regular expression.

Screen reader:

- A pattern that matches many strings -- every number, every line starting with a dash. Control plus F3 finds with one; Control plus Shift plus R replaces.

### Step 21

result window.

Screen reader:

- A document EdSharp opens with the output of a compile, an AI answer or a summary; read-only until you say otherwise.

### Step 22

run.

Screen reader:

- Executing the current file by its extension. F5; Shift plus F5 runs the address at the cursor.

### Step 23

say status.

Screen reader:

- Whether the document has changed since it was saved; pressed twice, its encoding. Alt plus Z.

### Step 24

session log.

Screen reader:

- What EdSharp did this session, kept under your local application data. Control plus F12 copies its path for a report.

### Step 25

snippet.

Screen reader:

- A saved piece of text or code, pasted or run with Alt plus V, saved with Alt plus S.

### Step 26

yield.

Screen reader:

- The size of the text: characters, words and lines, for all of it or the selection. Alt plus Y.

### Step 27

Twenty-six terms. The guide, F1, has each of them in context.

**Something to try:** Pick three terms you did not know and find each one in EdSharp.

## 10 - Conclusion

What to carry away from the walks: four principles, one thing kept from each task walk, three habits, and where to begin. I say the idea; the reader says the key.

**Before you start:** Nothing is needed.

### Step 1

Four things to carry away, each with its key. A key is named for a word of its command, so it can be guessed; and when it cannot, one list names every command and its key.

### Step 2

The Alternate Menu.

Screen reader:

- Alt plus F10

### Step 3

Alt says, Control does, Shift widens. Where am I?

Screen reader:

- Alt plus A

### Step 4

Has the document changed since I saved it?

Screen reader:

- Alt plus Z

### Step 5

A document in any format becomes text you can read, and your text becomes any format.

Screen reader:

- Control plus Shift plus O

### Step 6

And the AI is on your own machine; nothing leaves it.

Screen reader:

- Shift plus F12

### Step 7

One thing kept from each task walk. Walk four: open, change, save -- and F7 for the spelling before anyone else reads it.

Screen reader:

- F7

### Step 8

Walk five: a place is found by a word, a line or a bookmark, and a bookmark is kept with the file.

Screen reader:

- Control plus K

### Step 9

Walk six: Word, PDF and slides open as text; Markdown goes out as a web page; the preview is the check before you send.

Screen reader:

- Alt plus Shift plus E

### Step 10

Walk seven: choose the compiler once, compile, land on the error, move by blocks.

Screen reader:

- Control plus F5

### Step 11

Walk eight: ask the document a question, and translate it, on your own computer.

Screen reader:

- Alt plus Shift plus F7

### Step 12

Three habits that make the rest easy. Arrow a menu once a week: each item says its key. Open the Alternate Menu when a key is unknown and type part of the command's name. And Control plus F1, then press a key, to hear what it does without doing it.

Screen reader:

- Key Describer

### Step 13

Where to begin. Start with the file you work in most. Open it, change a line, save it, and let the keys come as you need them; the guide and the Key Describer are there for the ones you do not know yet.

### Step 14

When EdSharp is yours, make it yours further: Configuration Options for the spell checker and the voice of its messages, snippets for the text you type again and again, a compiler for each language you write.

### Step 15

If something goes wrong, Control plus F12 copies the session log's path; that file with a line about what you were doing, sent to the project page on GitHub, is the fastest way to a fix.

### Step 16

Three things people do with EdSharp every day, each in one breath. Read a document someone sent: Control plus Shift plus O, and it is text. Fix a program: Control plus F5, land on the error, fix it, Control plus F5. Write for the web: Markdown, Control plus F9 to check, Alt plus Shift plus E to export.

### Step 17

And three things the other Homer programs share with it, so the second program costs nothing to learn: the dialog that works one way, the Say keys that change nothing, and the Key Describer that answers any key without doing it.

### Step 18

Where the walks stop and the guide begins: scripting EdSharp in JScript, the Markdown code-block runner, the translator's quick start, the slide presenter's, the screen reader script writer's. F1 has all of it, and Control plus Shift plus F1 picks the quick start for your role.

**Something to try:** Open the file you work in most, change one line, save it; then do one thing from each task walk in it.

## 11 - More Information

Where the rest is: the guide and history from inside EdSharp, the documents, the GitHub page, updates, the session log, and the other Homer Tools.

**Before you start:** Nothing is needed.

### Step 1

F1 opens the guide, the whole of EdSharp in one document, from inside the program. Shift plus F1 opens the history of changes. Alt plus F1 says the version. Control plus Shift plus F1 opens the quick starts by role.

### Step 2

The ReadMe is the short start; the guide is the reference; Hotkeys lists every key three ways; the FAQ holds the questions people actually ask. All are in the help folder of the installation, and on the project's GitHub page.

### Step 3

F11 checks for a newer version and offers to install it. The project is at github dot com, slash JamalMazrui, slash EdSharp. Control plus F12 copies this session's log to the clipboard, ready to attach to a report of a problem.

### Step 4

EdSharp is one of the Homer Tools, free programs for working by ear: EdSharp for text, FileDir for files, DbDo for data. They share their keys and their player, so learning one is most of learning the next.

**Something to try:** Press F1 and read the first section of the guide.

<!-- walkthrough ends -->
