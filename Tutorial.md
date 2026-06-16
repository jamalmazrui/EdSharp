# Using EdSharp

A hands-on tutorial for getting started with the EdSharp text editor.

This tutorial is adapted, with thanks, from the *Using TextPal* tutorial originally written by Jim Homme for TextPal, a predecessor of EdSharp that contributed many of its ideas. The text here has been revised to match EdSharp's terminology, concepts, and key bindings. For the complete reference, see the EdSharp User Guide (press F1 in EdSharp).

## Contents

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

## Introduction

EdSharp is a full-featured, friendly, powerful, open-source text editor. It uses a standard Windows interface that supports multiple document windows, and it seeks to optimize efficiency for screen reader users by automatically speaking relevant information.

EdSharp works like Notepad, so you can begin using it with the same commands you already know. It adds many more commands that generally involve a modifier key such as Shift, Control, or Alt combined with a letter that begins the name of the command. Most commands provide enhanced speech compared to default screen reader output; for example, EdSharp reads the current line after a command completes. In general, EdSharp does not limit the size of a file or the number of open files, and it includes commands that make it friendly for programmers who want to use it to develop software.

## Installation

The installation program for EdSharp is called EdSharp_Setup.exe. When you run it, it does the following:

- Prompts for a program folder. The default is `C:\Program Files (x86)\EdSharp`.
- Creates a program group on the Windows Start menu with choices to launch EdSharp, read the documentation, or uninstall the program.
- Offers to set or clear an association between EdSharp and files with a particular extension, such as `.txt` or `.ini`. Binary formats such as `.pdf` or `.pptx` may also be associated with EdSharp so they are automatically converted to text when opened from Windows Explorer.
- Presents a checkbox for an optional set of JAWS scripts that fine-tune the EdSharp speech interface. If you install the scripts and later prefer default JAWS behavior, you can disable them from the JAWS "Manage Application Settings" dialog.
- Presents a checkbox that sets Alt+Control+E as a system-wide hot key for the EdSharp shortcut placed on the Windows desktop. If that hot key conflicts with another shortcut, select either shortcut on the desktop, press Alt+Enter for its properties, and change or clear the hot key.
- Presents a checkbox to open the manual in your default web browser after installation.

You can safely install new versions of EdSharp over previous versions; your settings, Favorites and Recent Files lists, and bookmarks are preserved. To update an existing installation, just press F11 (Elevate Version) and EdSharp downloads and installs the current version from the author's site. To check your version at any time, use the About command, Alt+F1; the History of Changes command, Shift+F1, summarizes fixes and improvements over time.

## Quick Start

If you are impatient to get going, here are some things you can do to get up and running.

- Launch EdSharp from the Start menu, the desktop shortcut, or its desktop hot key Alt+Control+E.
- Press F1 to open the full documentation in your default web browser.
- Press Alt+Shift+H for the Hotkey Summary, an alphabetically sorted list of hot keys placed in a new EdSharp window.
- Press Control+F1 for the Key Describer, then press any key to hear which command it runs.
- Press Alt+F10 for the Alternate Menu, an alphabetical list of every available command.
- Explore the regular menus to get a feel for the large number of familiar and new commands available through hot keys.

## Common Hot Keys

Here is a list of hot keys you will use often. Where EdSharp differs from what you may remember from other editors, the EdSharp binding is the one shown.

### Working With Files on Disk

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

### Navigating in the Current File

- Arrow keys, Page Up, and Page Down = Move through text. Add Shift to select.
- Alt+Down and Alt+Up = Next and prior sentence (EdSharp speaks the sentence).
- Control+Down and Control+Up = Next and prior paragraph (EdSharp speaks the paragraph).
- Control+K = Set a bookmark. Alt+K = Go to a bookmark. Control+Shift+K = Clear the bookmark (the cursor must be on the exact marked character).
- Control+J = Jump to a line number. Alt+J = Jump again. You may add a `+` or `-` before the number to jump that many lines forward or backward from the current line.
- Control+G = Go to a percentage point in the file. Alt+G = Go to that percentage again.

### Using the Clipboard

- F8 = Start Selection. Move the cursor to one character past the end of the text you want, then press Shift+F8 to Complete Selection.
- Shift plus the arrow keys = Select by character, word, or line in the usual way.
- Control+C = Copy. Control+X = Cut. Control+V = Paste. Each works on the current line if there is no selection.
- Alt+C = Copy and append to the clipboard. Alt+X = Cut and append to the clipboard.
- Alt+Apostrophe = Speak the current clipboard text.

### Changing Character Case

- Control+U = Upper case the current character or selection.
- Control+Shift+U = Lower case the current character or selection.
- Alt+U = Proper case (capitalize the first letter of each word, lower the rest).
- Alt+Shift+U = Swap case (invert upper and lower case).

### Searching and Replacing

- Control+F = Find forward. Control+Shift+F = Find backward.
- F3 = Find again forward. Shift+F3 = Find again backward.
- Alt+F3 = Find the chunk at the cursor, or the selected text, forward. Alt+Shift+F3 = the same, backward.
- Control+F3 = Find forward with a regular expression. Control+Shift+F3 = the same, backward.
- Control+R = Replace throughout all or selected text.
- Control+Shift+R = Replace using a regular expression.

### Working With Sections

EdSharp lets you divide a document into named sections, which is the basis of the address-book exercise later in this tutorial.

- Control+Enter = Insert a Section Break.
- Control+PageDown and Control+PageUp = Move to the next and prior section.
- F6 = Go to Section: jump from a table-of-contents line to its section. Shift+F6 = Go to Contents: jump from a section back to its table-of-contents line.
- Alt+T = Speak the current section's topic (its first line).
- Alt+Shift+T = Text Contents: build a table of contents from the first line of each section.
- Control+F6 = Search for a topic by name. Alt+F6 = Search for that topic again.

### Spell Checker and Thesaurus

- F7 = Spell check.
- Shift+F7 = Thesaurus.

### Software Development

- Control+Space = Select Chunk. Good for grabbing a function call or header in one step. Shift+Backspace speaks the current chunk.
- Shift+Enter = Insert a new line indented to the current line's level. Alt+Shift+Enter does the same but opens the line above.
- Tab = Indent the current or selected lines one level. Shift+Tab = Outdent them one level.
- Alt+I = Speak the indentation level of the current line.
- Alt+Shift+I = Indent Mode: toggle the spoken indentation alert on and off.
- Control+Q = Quote the selected text or whole document with a prefix string. You can set the prefix to match the comment string of your programming language to create comments, or use it for quoting email.
- Alt+PageDown and Alt+PageUp = Next and prior part: move to the next or previous function, method, or class, using the current compiler's NavigatePart pattern.
- Control+B and Control+Shift+B = Next and prior block: move by indentation, to the next or previous less-indented line.
- Control+F5 = Compile the current file, speak the output, and jump to the first error. With no compiler configured, a `.cs` file is compiled with the latest available .NET C# compiler automatically. Control+Shift+F5 picks or configures a compiler.

## The EdSharp Window

This section briefly explains the parts of the program you use most often: the title bar, the menu bar, the document area, and the status bar.

The title bar shows the name of the application and, in brackets, the document you are currently editing. The menu bar contains categories of commands you use throughout your work. The document area contains the document you are editing. The status bar reports information about what is happening as you work.

## Working With Files

This section explains how to create, open, and save files.

**Creating a new file.** When EdSharp first opens, it presents a blank document, ready for typing. To create another new file at any time, press Control+N; New is also the first choice on the File menu, Alt+F. To start a file from whatever is currently on the clipboard, press Control+Shift+N: EdSharp opens a new window, drops in the clipboard text, and places the cursor at the start.

**Opening files.** The simplest way to open a file is Control+O, which presents the standard Windows Open dialog showing the current folder. When you open a file this way, EdSharp converts several formats to plain text automatically: Microsoft Word, Excel, PowerPoint, PDF, Rich Text Format, and HTML. To open a file without conversion, use Control+Shift+O (Open Other Format); this is what you want for editing HTML source or importing an `.rtf` file with its formatting intact.

If you are editing a file and want to return to the version on disk, press Alt+O (Open Again). This only helps if you have not already saved your changes, since saving replaces the disk copy.

EdSharp keeps a Recent Files list, reached with Alt+R: pick a file and press Enter to load it. The Favorites list, Alt+L, stores files you return to often, such as a reference manual for a programming project, so you can open them without navigating the Open dialog. To add the file you are editing to Favorites, press Control+L while editing it. EdSharp does not convert files opened from the Recent Files or Favorites lists, so you can re-open a file in its native format; this is most useful for HTML files.

The File Find command, Alt+Shift+F, locates a file in the current folder from a text string: type the string and press Enter, then choose from the resulting list of matching files and press Enter to open it.

**Saving files.** Control+S is the key you will use most often. The first time you save a new file, EdSharp opens the Save As dialog so you can name it; afterward it simply writes the current version over the disk copy. Control+Shift+S (Save As) saves under a different name and switches the document window to the new file. Alt+Shift+S (Save Copy) writes a separate backup copy and leaves you working on the original, which is handy when you want to preserve a snapshot. For more file commands, explore the File menu, Alt+F.

## Working With Text

Besides entering and reviewing text, you will spend a lot of time manipulating it. This section covers the most common and useful operations.

**Selecting text.** Besides the usual Shift-with-navigation method, EdSharp offers two efficient techniques. The first is a two-key selection: place the cursor where the selection should start and press F8; move the cursor one character past the end of the text you want and press Shift+F8 to complete the selection. This is faster than holding Shift, because you do not have to keep a key down or wait for your screen reader to speak as you go. The second technique is Select Chunk, Control+Space, which is excellent for programmers: place the cursor in a function call and press Control+Space to select the entire call at once.

**Changing case.** Use Control+U for upper case, Control+Shift+U for lower case, Alt+U for proper case (the first letter of each word), and Alt+Shift+U to swap (invert) the case of the selection. These commands are also at the bottom of the Misc menu, Alt+M.

**Working with the clipboard.** You already know Control+X to cut and Control+C to copy a selection. EdSharp makes these more efficient: with no selection, Control+X cuts the current line and Control+C copies it, so you can skip selecting. Using Alt+X and Alt+C instead appends the cut or copied text to whatever is already on the clipboard, and these also work on the current line. You will find these and many more commands on the Edit menu, Alt+E.

**Navigating through text.** As you read, you will reach most often for next and prior paragraph (Control+Down and Control+Up), next and prior sentence (Alt+Down and Alt+Up), set and go to bookmark (Control+K and Alt+K), and find forward and backward (Control+F and Control+Shift+F, repeated with F3 and Shift+F3). The remaining navigation commands are on the Navigate menu, Alt+N.

## Getting Useful Information

EdSharp has a set of commands for learning about the document you are working on; most are on the Query menu, Alt+Q.

- Alt+Apostrophe speaks what is currently on the clipboard.
- Alt+A (Address) speaks the line and column of the cursor, followed by its percentage position from the top of the file.
- Alt+Y (Yield) reports how many characters, words, and lines the file, or the selection, contains.
- Alt+Z (Status) tells you whether the document has been modified from the version on disk; press it again to hear the character encoding.

## Creating an Address Book

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

## Conclusion

This tutorial has covered the commands you will use most often in day-to-day work with text, and it has shown how the section feature lets you build and navigate a structured document such as an address book and convert it to HTML. For the complete set of commands and features, including programming, math, word processing, and scripting, press F1 in EdSharp to open the full User Guide. EdSharp's author welcomes feedback and contributions toward its continued improvement.
