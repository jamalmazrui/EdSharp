# Fruit Basket on the Web

The same fruit basket program, written four ways for the browser. Each
one obeys the specification exactly: a fruit field with an Add button, a
basket list with a Delete button, a plain message when there is nothing
to add or delete, and focus that lands where the next action starts.

Reading them one after another is the point. The specification never
changes, so what differs between them is only how each framework thinks
about state, and how much ceremony it costs to keep a screen reader
happy.

## The programs

- **fruitBasketReact.htm** -- React 18. React from its official content
  delivery network, with Babel compiling the JSX in the browser, so
  there is no build step and nothing is installed.
- **fruitBasketVue.htm** -- Vue 3, in the official build that compiles
  templates in the browser. Nothing is installed.
- **fruitBasketWebComponent.htm** -- no framework at all: a custom
  element, which is the standards answer to the same problem. Nothing is
  downloaded either.
- **..\fruitBasket.htm** with **..\fruitBasket.js** -- plain HTML and
  JavaScript, in the folder above.

Frameworks that cannot be tried without a full project scaffold are left
out on purpose. Svelte and Angular both compile ahead of time and expect
a generated project with a package manifest and dozens of installed
packages, which is more ceremony than a fruit basket deserves; the web
component above shows the same ideas with nothing installed at all.

## Running them

Double-clicking a page usually works. If your browser objects to a page
loading its neighbours from disk, or you would rather serve them
properly, use the little server included here. It needs only Node
itself -- no npm install, no packages:

    node serveFruitBasket.js

It prints an address, opens your browser there, and lists the samples.
To choose a different port:

    node serveFruitBasket.js -port 8080

Press Control+C to stop it. Node is offered as a checkbox on the last
page of the EdSharp installer, in the same way Python is.

## What to listen for

Every version does these deliberately, and they are what make the
difference with a screen reader:

- Each label is tied to its field, so moving to the field says its name.
- The basket is a real list control, so arrow keys walk it and each
  item is announced with its position.
- Enter in the field adds a fruit; Delete inside the list removes one.
- Messages go to a live region, so they are spoken where you are
  instead of interrupting with an alert box you must dismiss.
- After adding, focus returns to the field; after deleting, to the list,
  with the neighbouring fruit selected so the list still has a place to
  speak from.
- Counts are worded to match: one fruit, not "1 fruits".

All are written in Camel Type, except where a framework insists
otherwise: React requires a component name to begin with a capital
letter, and a custom element's tag must contain a hyphen. Both are noted
in the code where they occur.
