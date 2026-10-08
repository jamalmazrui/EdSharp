---
title: Forms described by a file
lang: en
---

# Forms described by a file

A form here is a text file, not code. One section names the form, and every
section after it is one control. A program in any language that can write a
text file can put up a real Windows dialog and read back what the user chose.

This is IniForm, which Jamal Mazrui wrote in PowerBASIC between 2005 and
2015, brought forward into Lbc. The idea was unusual then and still is: a
single free program that turned a settings file into an accessible dialog.

## What changed, and why

**The file is .inix rather than .ini.** In the old format the items of a list
had to be crammed onto one line separated by vertical bars, and anything long
needed a second file with the text in it. An .inix value may span lines, so a
list reads as a list:

    [Address Type]
    control=list
    range=
      home
      work
      vacation
      other

Vertical bars still work, so every old definition still runs.

Indent those items if you like: the items of a range are trimmed one by one,
so the indentation is free. A memo is different, because every line of it is
content and its leading spaces may matter. .inix reads a multi-line value verbatim, from the start of its first line
to the end of its last, and trims nothing at either end; what to do with it
afterwards is the reading program's own business. So a memo whose text
should not be indented is written flush left, or between fences, which also
lets it hold lines that would otherwise look like keys:

    [Comments]
    control=memo
    value=`
    These are miscellaneous notes.
    This is a new line.
    `

**The dialog is built from ordinary Windows controls.** The old program built
Windows dialog templates by hand, in dialog units, which is why it could only
ever be 32-bit. Every control here is a normal WinForms control placed by
Lbc, so the form is whatever the program using it is. On a current build that
is 64-bit.

**Both layouts are available.** Set the form's layout to stack or to band.

*Stack* is Lbc's own and the default: one control to a row, so that reading
order, tab order and visual order are the same order. `align` is accepted
and ignored, since it has nothing to act on.

*Band* is the old arrangement, kept whole. Controls sit side by side where
`align=r` says so, a column of labels lines its colons up, and a definition
written for IniForm looks as it did. The arithmetic was transcribed from the
original and checked against the layout IniForm itself recorded for the
Customer Information sample: every position, every size, and the form
dimensions come out identical.

The trade is worth stating plainly. Band packs more into less space and
lines things up; in return the eye and the tab key may take different paths
across a row of controls. That is why stack is the default and band is there
when a form wants it.

## The control types

Nine came from IniForm: label, button, check, radio, list, multi, edit, memo
and status. Four are new:

- **combo** — a box you may type in, with a list beside it of the values
  already known. The value starts at whatever `value` says, in or out of the
  list, so a default nobody has used before is fine. The list is sorted
  alphabetically with upper and lower case treated alike.
- **droplist** — a drop-down offering the listed values and nothing else,
  which suits a long set of choices better than a list box sized for its
  widest item.
- **spin** — a number with a range, set by `min`, `max` and `step`.
- **heading** — a divider and a caption, for grouping fields.
- **separator** — the divider alone.

**password** was a `misc` value before and is a control type now, since that
is how people think of it. It still works either way.

## What a control section can say

- **caption** — the displayed name, when it differs from the section name.
- **control** — the type. An edit box when not given, as before.
- **help** — longer text, used as the tip when there is no tip.
- **min**, **max**, **step** — the bounds of a spin control.
- **misc** — `password`, `readonly`, `sort` or `nolabel`, separated by
  vertical bars.
- **range** — the items of a list, combo, droplist or multi.
- **selection** — which of them start selected, by position or by name.
- **tip** — the line shown in the status bar when the control has focus.
- **value** — what the control holds to begin with.

`nolabel` drops the visible label for a control that fills the form, but the
control is still named for a screen reader. An unlabelled control that
announces nothing is not a saving.

## What comes back

A section called Results, one key per control, written beside the input file
with `_input` changed to `_output`. A check box answers 1 or 0, as it always
did. A multi-selection answer is one item per line, which is what .inix made
possible and .ini did not.

Nothing is written when the form is cancelled. A cancelled form should not
leave answers lying about to be mistaken for real ones.

## Two things deliberately left behind

The old `input.txt` companion file, which existed only because .ini values
could not span lines or hold a vertical bar. An .inix value can do both, so
the companion has nothing left to do.

The per-control layout attributes — left, top, width, height, style, extend
— and the `output=all` mode that wrote them back. The arithmetic that
produced those numbers is kept, in band layout; what is not kept is the
ability to override an individual number by hand. Style and extend were
Windows bit flags for controls built from dialog templates, and the controls
here are ordinary WinForms ones.

One calculation was not carried forward, and it is worth saying why. A
band's shared button width — the width every button in a band takes, so that
they match — was written to the band but read from an array indexed by
control, so the maximum was taken against an unrelated and usually empty
slot. The effect was that a band took its *last* control's width rather than
its widest. With OK before Cancel that is the right answer by luck; with
Cancel before OK it clips the wider caption. The documented intent was
plain, so the intent is what is implemented.

## The examples

- **LogIn_input.inix** — the smallest useful form: two fields, two buttons.
- **CustomerInformation_input.inix** — the 2015 sample, carried forward: edit
  boxes, check boxes, a radio group, a list, a sorted multi-selection list
  and a memo.
- **NewControls_input.inix** — the types Lbc adds: headings, a combo and a
  spin control.
