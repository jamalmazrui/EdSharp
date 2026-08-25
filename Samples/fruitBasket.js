// fruitBasket.js -- the fruit basket program in standard JavaScript,
// paired with fruitBasket.htm.
//
// The fruit basket specification, from the listserv project of blind
// programmers: a fruit field with an Add button, a basket list with a
// Delete button, a plain message when there is nothing to add or
// delete, and focus that lands where the next action starts.
//
// The web version keeps the same manners as the desktop ones. The list
// is a real select element, so arrow keys walk it and every screen
// reader announces the item and the position without help. Messages go
// to a live region rather than to an alert box, so they are spoken
// where they happen and the keyboard never leaves the form.
//
// Written in Camel Type: Hungarian prefixes, lower camel case, one
// require or import per line, double quotes, functions rather than
// subprocedures.

"use strict";

var basketSelect, fruitInput, statusRegion;

// Say something to the person: the live region speaks it, and the text
// stays on screen for a sighted reader.
function sayStatus(sMessage) {
    if (statusRegion) statusRegion.textContent = sMessage;
}

// Add the field's fruit to the basket, select it, and return to the
// field ready for the next one.
function addFruit() {
    var optionFruit, sFruit;
    sFruit = fruitInput.value.trim();
    if (sFruit === "") {
        sayStatus("No fruit to add.");
        fruitInput.focus();
        return false;
    }
    optionFruit = document.createElement("option");
    optionFruit.textContent = sFruit;
    basketSelect.appendChild(optionFruit);
    basketSelect.selectedIndex = basketSelect.options.length - 1;
    fruitInput.value = "";
    sayStatus(sFruit + " added. " + countBasket());
    fruitInput.focus();
    return true;
}

// Remove the selected fruit, then select its neighbor so the list still
// has a place to speak from.
function deleteFruit() {
    var iFruit, sFruit;
    iFruit = basketSelect.selectedIndex;
    if (iFruit === -1) {
        sayStatus("No fruit to delete.");
        basketSelect.focus();
        return false;
    }
    sFruit = basketSelect.options[iFruit].textContent;
    basketSelect.remove(iFruit);
    if (iFruit > basketSelect.options.length - 1) iFruit = basketSelect.options.length - 1;
    if (iFruit >= 0) basketSelect.selectedIndex = iFruit;
    sayStatus(sFruit + " deleted. " + countBasket());
    basketSelect.focus();
    return true;
}

// How full the basket is, worded so one fruit is not "1 fruits".
function countBasket() {
    var iCount;
    iCount = basketSelect.options.length;
    if (iCount === 0) return "The basket is empty.";
    if (iCount === 1) return "1 fruit in the basket.";
    return iCount + " fruits in the basket.";
}

// Wire the page up once it exists.
function startFruitBasket() {
    basketSelect = document.getElementById("basket");
    fruitInput = document.getElementById("fruit");
    statusRegion = document.getElementById("status");

    document.getElementById("add").addEventListener("click", addFruit);
    document.getElementById("delete").addEventListener("click", deleteFruit);

    // Enter in the field adds, as the default button does on the desktop.
    fruitInput.addEventListener("keydown", function (oEvent) {
        if (oEvent.key === "Enter") { oEvent.preventDefault(); addFruit(); }
    });

    // Delete from the list itself, where the hand already is.
    basketSelect.addEventListener("keydown", function (oEvent) {
        if (oEvent.key === "Delete") { oEvent.preventDefault(); deleteFruit(); }
    });

    sayStatus("Ready. " + countBasket());
    fruitInput.focus();
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", startFruitBasket);
else startFruitBasket();
