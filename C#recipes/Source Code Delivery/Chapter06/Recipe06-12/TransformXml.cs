using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml.Xsl;
using System.Xml;

namespace Apress.VisualCSharpRecipes.Chapter06
{
    public partial class TransformXml : Form
    {
        public TransformXml()
        {
            InitializeComponent();
        }

        private void TransformXml_Load(object sender, EventArgs e)
        {
            XslCompiledTransform transform = new XslCompiledTransform();

            // Load the XSL stylesheet.
            transform.Load(@"..\..\orders.xslt");

            // Transform orders.xml into orders.html using orders.xslt.
            transform.Transform(@"..\..\orders.xml", @"..\..\orders.html");

            webBrowser1.Navigate(Application.StartupPath + @"\..\..\orders.html");
        }
    }
}