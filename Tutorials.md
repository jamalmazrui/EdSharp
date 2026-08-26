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
