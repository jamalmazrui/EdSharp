# Camel Type: JAWS Script Coding Guidelines

Camel Type is a coding style optimized for screen reader productivity and
systematic readability. This document adapts the Camel Type rules to the JAWS
scripting language as used in EdSharp's `.jss` script files and the related
JAWS settings files.

JAWS Script, like VBScript, is **case-insensitive**, so these guidelines follow
the same shape as the Camel Type rules for VBScript: lower camel case is a
discipline for readability rather than a requirement of the language. Where
JAWS Script differs from VBScript, the JAWS-specific rule is called out below.

---

## 1. File Family

A JAWS application script set is a family of files that share a base name (for
EdSharp, `EdSharp`) and differ by extension. Preserve the whole family; each
plays a role:

- `.jss` — script source (compiled to `.jsb`).
- `.jsh` — header file, brought in with `Include`.
- `.jsb` — compiled binary, brought in with `Use` or produced by the compiler.
- `.jkm` — key map: binds keystrokes to script names.
- `.jcf` — configuration file (script-specific options).
- `.jsd` — documentation strings shown by the Key Describer / help.

There is no equivalent of `Option Explicit`; instead, the discipline is to
declare every variable with `Var` (see section 5) before use.

---

## 2. Variable and Argument Naming

Use the same Hungarian prefixes as the VBScript guidelines to indicate type:

- `a` — array
- `b` — boolean
- `bin` — binary buffer
- `dt` — date-time
- `f` — file object
- `h` — handle (window handle, or a JAWS object/window handle)
- `i` — integer
- `l` — list
- `n` — real number
- `s` — string
- `d` — dictionary-like object
- `o` — generic object (any COM object obtained via the JAWS object functions)
- `v` — variant / unknown or mixed type

For a named COM object obtained in a script, use a prefix that reflects the
object class or a well-known abbreviation (for example `re` for a regular
expression object, `fso` for a file-system object). If only one instance of
that object type is in scope, the abbreviation is the entire variable name.

---

## 3. Constant Naming

Constants use the **same lower camel case naming convention as variables**,
including the Hungarian type prefix, and are distinguished only by being
declared with `Const`. Examples: `Const sDefaultVoice = "Eloquence"`,
`Const iSpeedStep = 5`.

JAWS built-in constants (for example the message-box and speech constants
defined in `HJConst.jsh`, such as `OT_LINE`, `MB_OKCANCEL`, `True`, `False`)
must always be used by their built-in name, never by their underlying numeric
value.

---

## 4. Capitalization

Use **lower camel case** for all custom names: variables, constants, `Function`
names, and `Script` names. JAWS Script is case-insensitive, but Camel Type
enforces lower camel case as a discipline for readability.

When calling JAWS built-in functions (for example `SayString`, `SayLine`,
`GetCursorRow`, `Speak`) or referencing built-ins from `HJConst.jsh`, use the
capitalization that the JAWS documentation specifies; these are not custom
names.

Note on key maps: because JAWS Script is case-insensitive, a `.jkm` binding such
as `Control+C=copySelectedTextToClipboard` resolves to a script declared as
`Script copySelectedTextToClipboard()` regardless of case, so re-casing existing
script names to lower camel case does not break existing key maps.

---

## 5. Variable and Constant Declarations

- Declare constants with `Const` and variables with `Var`.
- Declare them at the **top of the script file** (global scope, after `Include`
  and `Use` lines) or at the **top of each `Script` / `Function`**, before any
  logic.
- Group `Const` lines by type first, then `Var` lines by type.
- One `Const` or `Var` statement per type group, names listed **alphabetically**
  and **comma-separated** on a single line.
- Type groups appear in **alphabetical order** by prefix letter.

Example:

```
Const iSpeedStep = 5
Const sLogName = "Speech.log"

Var
    int iColumn, iRow,
    string sText, sTitle
```

(JAWS Script uses a single `Var` block with one type keyword per line; keep the
names on each type line alphabetical, and the type lines in alphabetical order
by prefix.)

---

## 6. Script Structure and Entry Point

A `.jss` file begins with its `Include` and `Use` lines, then any `Globals` /
`Const` declarations, then the routines. Unlike a standalone VBScript there is
no single entry point — JAWS invokes individual `Script` routines in response
to keystrokes (via the `.jkm`) and events. Keep the file's opening section
(includes, globals, constants) minimal and place the routines after it.

```
; EdSharp.jss -- JAWS scripts for EdSharp.exe
Include "HJConst.jsh"
Use "Homer.jsb"

Const iSpeedStep = 5

Script bottomOfFile()
    ; ...
EndScript
```

---

## 7. Routine Order

List all routines in **strict alphabetical order** by name, mixing `Script` and
`Function` routines together in that one order. There are no section dividers
grouping routines — alphabetical order is sufficient for navigation.

---

## 8. Scripts vs Functions

VBScript Camel Type says "define all routines as Functions." JAWS Script needs
one adaptation, because only a **`Script`** can be bound in a `.jkm` key map or
invoked as an event handler:

- Use **`Script`** for any routine that is bound to a keystroke or is a JAWS
  event handler. This is required by JAWS, not a style choice.
- Use **`Function`** for every helper routine that is called only from other
  script code. Prefer a `Function` (which can return a value) over a `Void`
  routine, and have it return a meaningful value even when the caller usually
  ignores it.
- Use **single-line `If`** syntax for simple one-consequence conditionals:

```
If sText == "" Then SayString ("Blank") EndIf
```

(JAWS Script requires the `EndIf` terminator; keep the whole
condition-and-consequence on one line when it is a single simple action.)

---

## 9. Loops

Prefer a collection-style iteration when the language and API allow it. JAWS
Script's looping constructs are `While ... EndWhile` and the counted
`For i = start To end ... EndFor`. Use the counted `For` only when the index
itself is needed; when simply walking a sequence, drive a `While` loop off a
cursor/iterator function rather than an explicit index where that reads more
clearly.

---

## 10. String Delimiters

JAWS Script string literals use **double quotes** only. Use `+` for string
concatenation.

```
sPath = sScriptDir + "\" + sName
```

---

## 11. Magic Numbers

Never use a literal numeric or string value in logic. Assign it to a named
`Const` and reference the constant. Always use the built-in named constants
from `HJConst.jsh` (for example `OT_SCREEN`, `VARIABLE_TYPE`) rather than their
numeric equivalents.

---

## 12. Object Creation

When a script needs a COM object, create it with the JAWS object function
(for example `CreateObjectEx`) using the full ProgID string, assign it to a
variable named with the object's prefix convention, and release it when done.

```
Const sReProgId = "VBScript.RegExp"
Var object oRe
Let oRe = CreateObjectEx (sReProgId, ...)
; ... use oRe ...
Let oRe = Null
```

---

## 13. Error Handling

JAWS Script does not provide structured `try`/`catch`. Guard operations that
can fail by checking the relevant status or return value immediately after the
call, and keep the guarded region as small as possible. Where a JAWS built-in
reports failure through a return code, test that return code on the next line
rather than assuming success.

---

## 14. Comments

Use the semicolon (`;`) line comment for ordinary remarks and `/* ... */` for
block comments. Lead each `Script` or `Function` with a short comment stating
what it does and, for a bound `Script`, the keystroke that invokes it (mirrored
in the `.jkm` and `.jsd`).
