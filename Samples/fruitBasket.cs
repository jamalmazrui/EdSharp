// fruitBasket.cs -- the fruit basket program in C# with Windows Forms.
//
// The fruit basket specification, from the listserv project of blind
// programmers: a window with a fruit field and an Add button, a basket
// list and a Delete button. Adding an empty field says so rather than
// adding nothing; deleting with nothing selected says so too; after
// each action the focus lands where the next action starts, which is
// what makes the program pleasant with a screen reader.
//
// Written in Camel Type: Hungarian prefixes on names, lower camel case
// for everything the language allows, functions rather than
// subprocedures, declarations grouped at the top of each function.
//
// Build with the C# compiler that ships with Windows:
//   csc.exe /nologo /target:winexe /platform:x64 fruitBasket.cs

using System;
using System.Drawing;
using System.Windows.Forms;

namespace FruitBasket {

public class fruitForm : Form {
Button btnAdd, btnDelete;
Label lblBasket, lblFruit;
ListBox lbBasket;
TextBox txtFruit;

public fruitForm() {
this.Text = "Fruit Basket";
this.StartPosition = FormStartPosition.CenterScreen;
this.ClientSize = new Size(460, 300);

// The layout is a table so the window can be resized and so the
// reading order matches the tab order: label, field, button on the
// first row; label, list, button on the second.
TableLayoutPanel tblLayout = new TableLayoutPanel();
tblLayout.Dock = DockStyle.Fill;
tblLayout.ColumnCount = 3;
tblLayout.RowCount = 2;
tblLayout.Padding = new Padding(12);
tblLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
tblLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
tblLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
tblLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
tblLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

lblFruit = new Label();
lblFruit.Text = "&Fruit:";
lblFruit.AutoSize = true;
lblFruit.Anchor = AnchorStyles.Left;

txtFruit = new TextBox();
txtFruit.Dock = DockStyle.Fill;
txtFruit.AccessibleName = "Fruit";

btnAdd = new Button();
btnAdd.Text = "&Add";
btnAdd.AutoSize = true;
btnAdd.Click += new EventHandler(this.addClick);
this.AcceptButton = btnAdd;

lblBasket = new Label();
lblBasket.Text = "&Basket:";
lblBasket.AutoSize = true;
lblBasket.Anchor = AnchorStyles.Left | AnchorStyles.Top;

lbBasket = new ListBox();
lbBasket.Dock = DockStyle.Fill;
lbBasket.AccessibleName = "Basket";
lbBasket.KeyDown += new KeyEventHandler(this.basketKeyDown);

btnDelete = new Button();
btnDelete.Text = "&Delete";
btnDelete.AutoSize = true;
btnDelete.Anchor = AnchorStyles.Top;
btnDelete.Click += new EventHandler(this.deleteClick);

tblLayout.Controls.Add(lblFruit, 0, 0);
tblLayout.Controls.Add(txtFruit, 1, 0);
tblLayout.Controls.Add(btnAdd, 2, 0);
tblLayout.Controls.Add(lblBasket, 0, 1);
tblLayout.Controls.Add(lbBasket, 1, 1);
tblLayout.Controls.Add(btnDelete, 2, 1);
this.Controls.Add(tblLayout);
} // fruitForm constructor

// Add the field's fruit to the basket, select it, and return to the
// field ready for the next one.
bool addFruit() {
string sFruit = txtFruit.Text.Trim();
if (sFruit.Length == 0) {
MessageBox.Show("No fruit to add.", "Alert");
txtFruit.Focus();
return false;
}
lbBasket.Items.Add(sFruit);
lbBasket.SelectedIndex = lbBasket.Items.Count - 1;
txtFruit.Text = "";
txtFruit.Focus();
return true;
} // addFruit function

// Remove the selected fruit, then select its neighbor so the list
// still has a place to speak from.
bool deleteFruit() {
int iFruit = lbBasket.SelectedIndex;
if (iFruit == -1) {
MessageBox.Show("No fruit to delete.", "Alert");
lbBasket.Focus();
return false;
}
lbBasket.Items.RemoveAt(iFruit);
if (iFruit > lbBasket.Items.Count - 1) iFruit = lbBasket.Items.Count - 1;
if (iFruit >= 0) lbBasket.SelectedIndex = iFruit;
lbBasket.Focus();
return true;
} // deleteFruit function

void addClick(object oSender, EventArgs oArgs) {
addFruit();
} // addClick handler

void deleteClick(object oSender, EventArgs oArgs) {
deleteFruit();
} // deleteClick handler

// Delete from the list itself, which is where a person's hand already
// is when they decide a fruit should go.
void basketKeyDown(object oSender, KeyEventArgs oArgs) {
if (oArgs.KeyCode == Keys.Delete) { deleteFruit(); oArgs.Handled = true; }
} // basketKeyDown handler

[STAThread]
public static void Main() {
Application.EnableVisualStyles();
Application.Run(new fruitForm());
} // Main function

} // fruitForm class

} // FruitBasket namespace
