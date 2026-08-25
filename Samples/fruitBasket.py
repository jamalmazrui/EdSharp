"""fruitBasket.py -- the fruit basket program in Python with wxPython.

The fruit basket specification, from the listserv project of blind
programmers: a fruit field with an Add button, a basket list with a
Delete button, a plain message when there is nothing to add or delete,
and focus that lands where the next action starts.

wxPython draws native Windows controls, so the field, list and buttons
are the same ones every other program uses and every screen reader
already knows them. Sizers place the controls, which keeps the reading
order and the tab order the same as the visual order -- the part that
matters most when the layout is heard rather than seen.

Written in Camel Type: Hungarian prefixes on names, lower camel case for
functions and variables, functions rather than subprocedures, built-in
modules imported before third-party ones.

Install the requirement once:

    python -m pip install wxPython

Then run:

    python fruitBasket.py
"""

import sys

import wx


class fruitFrame(wx.Frame):
    """The one window: fruit in, basket out."""

    def __init__(self):
        wx.Frame.__init__(self, None, title="Fruit Basket", size=(480, 320))
        panelMain = wx.Panel(self)

        self.lblFruit = wx.StaticText(panelMain, label="&Fruit:")
        self.txtFruit = wx.TextCtrl(panelMain, style=wx.TE_PROCESS_ENTER)
        self.btnAdd = wx.Button(panelMain, label="&Add")
        self.lblBasket = wx.StaticText(panelMain, label="&Basket:")
        self.lbBasket = wx.ListBox(panelMain, choices=[])
        self.btnDelete = wx.Button(panelMain, label="&Delete")
        self.lblStatus = wx.StaticText(panelMain, label="The basket is empty.")

        # Enter in the field adds, as the default button does elsewhere.
        self.btnAdd.SetDefault()
        self.btnAdd.Bind(wx.EVT_BUTTON, self.addClick)
        self.btnDelete.Bind(wx.EVT_BUTTON, self.deleteClick)
        self.txtFruit.Bind(wx.EVT_TEXT_ENTER, self.addClick)
        self.lbBasket.Bind(wx.EVT_KEY_DOWN, self.basketKeyDown)

        # Two rows of three, then a status line across the bottom.
        sizerGrid = wx.FlexGridSizer(rows=2, cols=3, vgap=8, hgap=8)
        sizerGrid.AddGrowableCol(1, 1)
        sizerGrid.AddGrowableRow(1, 1)
        sizerGrid.Add(self.lblFruit, 0, wx.ALIGN_CENTER_VERTICAL)
        sizerGrid.Add(self.txtFruit, 1, wx.EXPAND)
        sizerGrid.Add(self.btnAdd, 0)
        sizerGrid.Add(self.lblBasket, 0, wx.ALIGN_TOP)
        sizerGrid.Add(self.lbBasket, 1, wx.EXPAND)
        sizerGrid.Add(self.btnDelete, 0, wx.ALIGN_TOP)

        sizerMain = wx.BoxSizer(wx.VERTICAL)
        sizerMain.Add(sizerGrid, 1, wx.EXPAND | wx.ALL, 12)
        sizerMain.Add(self.lblStatus, 0, wx.LEFT | wx.BOTTOM, 12)
        panelMain.SetSizer(sizerMain)

        self.Centre()
        self.txtFruit.SetFocus()

    def sayStatus(self, sMessage):
        """Put a short message where both eye and screen reader find it."""
        self.lblStatus.SetLabel(sMessage)

    def countBasket(self):
        """How full the basket is, worded so one fruit is not "1 fruits"."""
        iCount = self.lbBasket.GetCount()
        if iCount == 0: return "The basket is empty."
        if iCount == 1: return "1 fruit in the basket."
        return str(iCount) + " fruits in the basket."

    def addFruit(self):
        """Add the field's fruit, select it, and return to the field."""
        sFruit = self.txtFruit.GetValue().strip()
        if not sFruit:
            wx.MessageBox("No fruit to add.", "Alert")
            self.txtFruit.SetFocus()
            return False
        self.lbBasket.Append(sFruit)
        self.lbBasket.SetSelection(self.lbBasket.GetCount() - 1)
        self.txtFruit.SetValue("")
        self.sayStatus(sFruit + " added. " + self.countBasket())
        self.txtFruit.SetFocus()
        return True

    def deleteFruit(self):
        """Remove the selected fruit and select its neighbor."""
        iFruit = self.lbBasket.GetSelection()
        if iFruit == wx.NOT_FOUND:
            wx.MessageBox("No fruit to delete.", "Alert")
            self.lbBasket.SetFocus()
            return False
        sFruit = self.lbBasket.GetString(iFruit)
        self.lbBasket.Delete(iFruit)
        if iFruit > self.lbBasket.GetCount() - 1: iFruit = self.lbBasket.GetCount() - 1
        if iFruit >= 0: self.lbBasket.SetSelection(iFruit)
        self.sayStatus(sFruit + " deleted. " + self.countBasket())
        self.lbBasket.SetFocus()
        return True

    def addClick(self, oEvent):
        self.addFruit()

    def deleteClick(self, oEvent):
        self.deleteFruit()

    def basketKeyDown(self, oEvent):
        """Delete from the list itself, where the hand already is."""
        if oEvent.GetKeyCode() == wx.WXK_DELETE: self.deleteFruit()
        else: oEvent.Skip()


def main():
    oApp = wx.App(False)
    frameFruit = fruitFrame()
    frameFruit.Show()
    oApp.MainLoop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
