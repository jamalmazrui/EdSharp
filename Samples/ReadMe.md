# Fruit Basket Samples

These are working programs that meet the fruit basket specification, a
small exercise from a listserv project of blind programmers: one window
with a fruit field and an Add button, a basket list and a Delete button.
Adding an empty field says so rather than adding nothing. Deleting with
nothing selected says so too. After each action the focus lands where
the next action starts.

The exercise is small on purpose. It is just large enough to show how a
language and its user interface library handle the things that decide
whether a program is pleasant with a screen reader: labels tied to their
fields, a real list control, a default button, keyboard access to every
action, and messages that arrive where you are rather than where the
mouse is.

## The programs

- **fruitBasket.cs** -- C# with Windows Forms. Build it with the C#
  compiler that ships with Windows, or just open it in EdSharp and press
  Control+F5:

      csc.exe /nologo /target:winexe /platform:x64 fruitBasket.cs

- **fruitBasket.js** with **fruitBasket.htm** -- standard HTML and
  JavaScript, no framework. Open the page in any browser. The basket is
  a real select element, so arrow keys walk it, and messages go to a
  live region so they are spoken without moving the keyboard.

- **fruitBasket.py** -- Python with wxPython, which draws native Windows
  controls. Install the library once, then run it:

      python -m pip install wxPython
      python fruitBasket.py

- **Web** -- a folder of browser versions: React, Vue, a framework-free
  custom element, and a small Node server for trying them. Its own
  ReadMe explains each.

## What they share

All three obey the same rules, so reading one after another shows what
changes between languages and what does not:

- The Add button is the default, so Enter in the field adds a fruit.
- Adding selects the new fruit, clears the field, and returns focus to
  the field.
- Deleting selects the neighbor of the fruit removed, so the list still
  has a place to speak from, and returns focus to the list.
- The Delete key works inside the list itself, where your hand already
  is when you decide a fruit should go.
- Counts are worded to match: one fruit, not "1 fruits".

All three are written in Camel Type, the coding style used throughout
Homer Tools: Hungarian prefixes on names, lower camel case wherever the
language allows it, functions rather than subprocedures, and
declarations grouped at the top.
