using System;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace anrControls.MarkdownDriver
{
    public class _default : Page
    {
        #region Page controls

        protected TextBox MarkdownSource;
        protected DropDownList OutputOptions;
        protected Button Convert;
        protected TextBox HtmlSource;
        protected PlaceHolder plhHtmlsource;
        protected Label Preview;
        protected DropDownList FilterOptions;
        protected PlaceHolder plhHtmlPreview;

        #endregion

        private void Convert_Click(object sender, EventArgs e)
        {
            Markdown    md = new Markdown ();
            SmartyPants sm = new SmartyPants();
            string      markdownSource = MarkdownSource.Text;
            string      htmlSource = null;
            
            
            switch (FilterOptions.SelectedIndex)
            {
                case 0:     htmlSource = md.Transform (markdownSource); break;
                case 1:     htmlSource = sm.Transform (markdownSource, ConversionMode.EducateDefault); break;
                default:    htmlSource = sm.Transform (md.Transform(markdownSource), ConversionMode.EducateDefault); break;
            }

            HtmlSource.Text = htmlSource;
            Preview.Text = htmlSource;

            switch (OutputOptions.SelectedIndex)
            {
                case 0: plhHtmlsource.Visible = true; plhHtmlPreview.Visible = false; break;
                case 1: plhHtmlsource.Visible = false; plhHtmlPreview.Visible = true; break;
                default: plhHtmlsource.Visible = true; plhHtmlPreview.Visible = true; break;
            }
        }

        #region Web Form Designer generated code

        protected override void OnInit (EventArgs e)
        {
            base.OnInit (e);
            Convert.Click += new EventHandler(Convert_Click);
        }
        #endregion
    }
}