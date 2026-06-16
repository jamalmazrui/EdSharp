import System
import System.Drawing
import System.Windows.Forms

var dlg = new Form()
dlg.Text = 'RichTextBox Example'
dlg.Width = 400
dlg.Hight = 400

var rtb = new RichTextBox
rtb.Dock = DockStyle.Fill
dlg.Controls.Add(rtb)
dlg.StartPosition = FormStartPosition.CenterScreen
dlg.ShowDialog()
