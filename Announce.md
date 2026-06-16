# EdSharp 5.0 beta

EdSharp is a free, open-source, accessible text and source-code editor for Windows, designed to be efficient for screen-reader users. Version 5.0 is the first 64-bit release and a broad modernization of the program. This is a beta: the core is in place and building cleanly, and feedback is welcome before the final release.

## Key improvements in 5.0

**Modern 64-bit foundation.** EdSharp is now a 64-bit application targeting the .NET Framework 4.8 and built with the current Roslyn C# compiler. `BuildEdSharp.cmd` compiles the program and retrieves everything it needs, so a developer can rebuild from source with one command.

**Layered speech.** Spoken feedback is delivered through whichever channel is available, in order: JAWS, NVDA, Windows UI Automation, and SAPI. This makes EdSharp's extra announcements work across screen readers without special configuration.

**Automatic character-encoding detection.** EdSharp detects a file's encoding from its byte-order mark and, when there is none, from its content, so files in legacy code pages or UTF-16 open correctly. New files default to UTF-8. Each document remembers its encoding for saving.

**A more capable configuration format.** Settings can now be expressed in the new `.inix` format, which preserves order and supports multi-line values, while remaining backward compatible with classic `.ini` files.

**A central command map.** A single internal command table now drives the Hotkey Summary (Alt+Shift+H), the Key Describer (Control+F1), and the Alternate Menu (Alt+F10), so the spoken key help stays consistent everywhere.

**Compiling that works out of the box.** Press Control+F5 on a C# file with no setup and EdSharp compiles it with the latest available .NET C# compiler and jumps to the first error. Compilers for other languages can be defined robustly in the `.inix` format, with regular expressions that locate the error line and column.

**Regexer-style transform jobs.** The Transform Files command (Alt+Equals) now reads a job as a Regexer `.inix` file: one named section per task, each with Find, Replace, Options, Extract, and Divider keys, and values that may span multiple lines.

**Modernized web download.** Downloading from a web page now negotiates current TLS, sends a normal browser User-Agent, follows redirects, and names files from the server's Content-Disposition header (guessing an extension from the content type when needed). Gathering links from a page no longer depends on Internet Explorer, which has been removed from current Windows.

**Bundled conversion tools, fetched on demand.** At build time EdSharp retrieves current portable builds of Pandoc, HTML Tidy, liblouis, Xpdf, and Artistic Style, and keeps the Pandoc configuration up to date, so import, export, and code-formatting features work without manual setup.

**Documentation.** The user guide has been refreshed, its reference links updated to current sources, and a new step-by-step Tutorial (adapted with thanks from Jim Homme's TextPal tutorial) is available from the Help menu and opens in your web browser.

## Feedback

Because this is a beta, please report any rough edges, with steps to reproduce where possible. The latest version can always be installed or updated from within EdSharp using the Elevate Version command, F11.
