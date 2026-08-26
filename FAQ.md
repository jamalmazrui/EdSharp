# EdSharp Questions and Answers

## Getting Started

### How do I start EdSharp?

Press Alt+Control+E from anywhere in Windows. If EdSharp is already running,
that key brings it forward rather than starting a second copy.

### What do I press first?

Control+F1 turns on Key Describer. Press any key and EdSharp says what it does
without doing it; press Control+F1 again to leave. Alt+F10 lists every command
alphabetically with its key, and Alt+Shift+H shows the whole table in a window
you can search.

### Do I need anything else installed?

No. EdSharp runs on the .NET Framework built into every modern Windows. The
optional pieces offered by the installer add features but nothing essential.

### Which boxes should I tick in the installer?

Python and the document tools are ticked already, and they are the ones that
make EdSharp itself better: rich PDF conversion and the thesaurus, about a
hundred and fifty megabytes together. Tick Ollama if you want the AI features,
and its two extra models if you want better translation and code help. Leave
Git and Node unless you already know you want them.

## Documents

### Can EdSharp open a Word document or a PDF?

Yes. Control+Shift+O, Open Other Format, converts it to text you can read and
search. A PDF keeps its headings, lists and tables, so you can navigate it by
block with Control+B rather than wading through a wall of text.

### Why does a PDF sometimes come out empty?

Because it is a scan of images rather than text. No converter can help with
that; it needs optical character recognition first. EdSharp says so rather
than handing you an empty file.

### How do I save as something else?

Alt+Shift+E, Export Format, writes the document out as a Word document, a web
page, a slide deck, plain text, and more. Control+T converts the document in
place.

### Can I make PowerPoint slides?

Yes. Write the talk in Markdown and export as pptx. Each second level heading
starts a slide, three hyphens on their own line start an untitled one, and a
notes section holds speaker notes that appear in Presenter View but not on the
slide. The slide presenter tutorial has the details.

## Writing

### F7 says no spell checker could be started. What now?

The message names the file it looked for. The dictionary lives in the
Dictionaries folder beside EdSharp, and the installer puts it there;
reinstalling restores it. If you have Microsoft Word and would rather use its
checker, set the SpellChecker option to Word with Configuration Options,
Alt+Shift+C.

### Does the spell checker learn words?

Yes. Add to Dictionary in the spell check dialog puts the word in a plain file
in your settings folder, which you can also edit by hand, and teaches Windows'
own dictionary at the same time.

### Why does the thesaurus give definitions?

Because a thesaurus without them is guesswork. The words are grouped by
meaning, so synonyms for light as in weight never mix with words about
illumination.

## Code

### How do I run the file I am editing?

Control+F5. If nothing happens, press Control+Shift+F5 first and choose the
compiler for your language: that one choice sets the command to run, the way
errors are read, the indentation, the comment prefix and the interactive
shell.

### The cursor lands on the wrong error.

It lands on the earliest error in the file, which is often not the first one
the compiler printed — compilers group warnings and errors as they please. Fix
it and press Control+F5 again.

### What is the snippet console?

Control+Shift+G opens a prompt where the editor window is in scope as frm and
its text box as rtb. Type an expression to see its value, or a statement to
change the document. Whatever works there works in a snippet, which is the
point. There are two, JScript and C#, chosen by which compiler you have
picked.

### Why does Python indent with four spaces when EdSharp prefers a tab?

Because a Python file is usually shared, and four spaces is what collaborators
expect. EdSharp's own default is one tab, and a file that already has
indentation of its own keeps it — the compiler setting only governs a file
with none.

## The AI Features

### Does anything leave my computer?

No. The models run on your own machine through Ollama. There is no account, no
key and no limit, and no document is sent anywhere.

### Do I need a graphics card?

Not for the small model, which answers promptly on any machine. The larger
translation and coding models are noticeably faster with one.

### F12 sometimes sends my document and sometimes does not.

It reads your wording. "Summarize this" plainly refers to the document, so the
document travels; "when is Thanksgiving this year" does not, so it goes alone
and answers in seconds. The status line says which happened. Shift+F12 always
sends the document.

### How good is the translation?

Good between the widely written languages — Spanish, French, German,
Portuguese, Italian — and weaker for languages with less material behind them.
The larger model, a checkbox in the installer, is better at all of them.

## Screen Readers

### Which screen readers work with EdSharp?

JAWS, NVDA and Narrator. EdSharp speaks through whichever is running, and the
installer offers to install JAWS scripts and an NVDA add-on that add EdSharp
specific behavior.

### Why is EdSharp quieter than it used to be?

Because your screen reader already announces window titles and the name and
value of whatever has focus, and hearing both was tiring. EdSharp now speaks
only what it alone knows: a count, a level, the text a command reached. If
somewhere feels too quiet, say so and it can be added back deliberately.

### Can I turn its speech off entirely?

Yes, with Extra Speech in the Miscellaneous menu. The messages still go to the
status bar and to a log you can read.

## Problems

### Something went wrong. What should I send?

Press Control+F12. That copies this session's log to the clipboard as a file,
ready to attach to a message. Add a sentence about what you were doing.

### Where does EdSharp keep its files?

Settings are in an EdSharp folder under your roaming application data; logs
are in an EdSharp logs folder under your local application data. The program
itself is in Program Files, and nothing personal is kept there.

### An installer checkbox says Reinstall when I have not installed it.

It means the tool is already on your computer, perhaps from another program.
Reinstall entries are never ticked by default, because there is nothing to
gain.

Jamal Mazrui
