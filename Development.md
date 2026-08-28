# EdSharp Development

How EdSharp is built, how it is put together, and the conventions it follows.
For using EdSharp, see the user guide; for learning it by role, the tutorials.

## Building

Clone the project and run **BuildEdSharp.cmd** from the project folder. It
needs nothing installed beyond Windows itself, though it uses more when it is
there.

The build does these things in order, writing everything to BuildEdSharp.log:

1. Works out the next version number from the release tags on the remote, and
   writes it into the installer script.
2. Fetches the spelling dictionary if it is missing.
3. Removes retired JAWS script bindings from the Scripts folder.
4. **Runs the source audit** and stops if any check fails. See below.
5. Fetches the pinned libraries from the package gallery when they are absent
   or the wrong version: the Markdown reader and writer, the HTML parser, the
   character-set detector, the spelling engine.
6. Compiles EdSharp.js into EdSharp.dll with the JScript compiler that ships
   with the framework, then compiles the C# sources into EdSharp.exe with the
   best C# compiler it can find — the Roslyn compiler from Visual Studio Build
   Tools if present, otherwise the one inside Windows.
7. Regenerates the .htm for every tracked Markdown file whose .htm is older.
8. Builds the installer if Inno Setup is present.

A build that fails says why in the log, and the log is the thing to send.

## The Source Audit

**auditEdSharp.cmd** checks what a compiler cannot. Each check exists because
something once broke:

- Braces balance in the C# sources, counted outside comments and strings.
- Brackets balance in every shipped PowerShell script.
- No two commands claim the same key.
- Every dialog's buttons have distinct access keys, with OK and Cancel
  deliberately having none, since Control+Enter and Escape serve them.
- Every command has a description, and each description names the key the
  code actually binds.
- The compiler table's regular expressions are legal expressions, and every
  compiler section defines its settings.
- Conversion entries name scripts that exist.
- The spell checking interfaces match the layouts Windows publishes, slot for
  slot.
- No public method exposes a type declared private in the same file.
- Types from namespaces the file does not import are written in full.
- Microsoft Office is reached only by the named fallbacks.

The build runs it before compiling. Run it by hand when a build dies without
writing a log — a script that will not parse never gets far enough to log
anything.

## How the Program Is Put Together

**EdSharp.cs** holds most of the program: the window, the commands, the
conversion and compile machinery. It is one large file, which is deliberate:
its author navigates by search and by structure rather than by file, and
splitting it would cost more than it saved.

**Lbc.cs** is the dialog toolkit — Layout by Code. Dialogs are built by adding
labelled fields in order rather than by a visual designer, which keeps the
reading order, the tab order and the visual order identical. Every dialog gets
Control+Enter to accept, Escape to cancel and F1 to describe its fields.

Its forms follow the Homer guidelines. Every control -- every tab stop, with a
label and its field counting as one -- carries a trigger letter marked by an
ampersand, so Alt and that letter reach it from anywhere in the form. The
letter must begin a word: the first word's initial by preference, the second
word's when that clashes, and a letter from inside a word almost never, the
rare exception being a strong mnemonic such as X for Export. When two controls
want the same letter, rename one. OK and Cancel take no letter, since
Control+Enter and Escape are their keys, and Help is added by the toolkit
holding H. The audit checks every one of these.

**Say.cs** speaks. It tries JAWS through its automation interface, then NVDA
through its controller library, then a system notification that Narrator
reads, and stops at the first that answers, so a message is never delivered
twice.

**Inix.cs** reads the settings files. **KeyMap.cs** maps keys to commands.
**Web.cs** fetches web pages.

**EdSharp.js** is compiled by the JScript compiler into EdSharp.dll and loaded
by reflection, never by a compile-time reference — the executable and the
library share a base name, and a compile-time reference would make loading
ambiguous. It runs snippets and hosts the JScript console.

## Settings

Three files, in increasing order of authority:

- **EdSharp.inix** in the program folder holds what EdSharp ships with: the
  conversion table, the compiler definitions, the defaults. Upgrades replace
  it, so nothing personal belongs here.
- **EdSharp.inix** in the data folder holds personal overrides in the same
  format. Upgrades never touch it.
- **EdSharp.ini** in the data folder holds settings changed through
  Configuration Options.

A compiler is a section named Compiler followed by its name, with one named
key per setting: the compile command, the pattern locating an error, the
output to abbreviate away, the pattern marking the start of a part, the
comment prefix, the default extension, the indentation the language uses, and
the interactive shell to open. Adding a compiler needs no code.

## Coding Style

EdSharp uses Camel Type: Hungarian prefixes on names — s for string, i for
integer, b for boolean, l for list, a for array, o for other objects — and
lower camel case wherever the language allows it. Functions rather than
subprocedures. Simple conditions on one line. Declarations grouped at the top.

The style optimizes for hearing code rather than seeing it: a prefix tells you
the type as the name is spoken, and one-line conditions spare a reader two
extra lines of arrow-keying for a two-word test.

The Samples folder holds the same program written several ways in this style,
which is the quickest way to absorb it.

## Documentation

Every Markdown file in the project root that git tracks has a matching .htm,
regenerated by the build whenever the Markdown is newer. Write the Markdown;
never edit the .htm.

Hotkeys.md is generated from the description table inside EdSharp.cs, which is
also what Key Describer and the Alternate Menu read. Changing a description
means changing that table, not the file.

## Releasing

Build, test, then **tagRelease**, which tags the version the build wrote into
the installer script and pushes it. Releases before v5.0.20 include
executables that cannot start and should not be installed.

Jamal Mazrui
