using System;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter06
{
    public partial class XmlTreeDisplay : Form
    {
        public XmlTreeDisplay()
        {
            InitializeComponent();
        }

        private void XmlTreeDisplay_Load(object sender, EventArgs e)
        {
            txtXmlFile.Text = Path.Combine(Application.StartupPath,
               @"..\..\ProductCatalog.xml");
        }

        private void cmdLoad_Click(object sender, System.EventArgs e)
        {
            // Clear the tree.
            treeXml.Nodes.Clear();

            // Load the XML Document
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(txtXmlFile.Text);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return;
            }

            // Populate the TreeView.
            ConvertXmlNodeToTreeNode(doc, treeXml.Nodes);

            // Expand all nodes.
            treeXml.Nodes[0].ExpandAll();
        }

        private void ConvertXmlNodeToTreeNode(XmlNode xmlNode, TreeNodeCollection treeNodes)
        {
            // Add a TreeNode node that represents this XmlNode.
            TreeNode newTreeNode = treeNodes.Add(xmlNode.Name);

            // Customize the TreeNode text based on the XmlNode
            // type and content.
            switch (xmlNode.NodeType)
            {
                case XmlNodeType.ProcessingInstruction:
                case XmlNodeType.XmlDeclaration:
                    newTreeNode.Text = "<?" + xmlNode.Name + " " + xmlNode.Value + "?>";
                    break;
                case XmlNodeType.Element:
                    newTreeNode.Text = "<" + xmlNode.Name + ">";
                    break;
                case XmlNodeType.Attribute:
                    newTreeNode.Text = "ATTRIBUTE: " + xmlNode.Name;
                    break;
                case XmlNodeType.Text:
                case XmlNodeType.CDATA:
                    newTreeNode.Text = xmlNode.Value;
                    break;
                case XmlNodeType.Comment:
                    newTreeNode.Text = "<!--" + xmlNode.Value + "-->";
                    break;
            }

            // Call this routine recursively for each attribute.
            // (XmlAttribute is a subclass of XmlNode.)
            if (xmlNode.Attributes != null)
            {
                foreach (XmlAttribute attribute in xmlNode.Attributes)
                {
                    ConvertXmlNodeToTreeNode(attribute, newTreeNode.Nodes);
                }
            }

            // Call this routine recursively for each child node.
            // Typically, this child node represents a nested element,
            // or element content.
            foreach (XmlNode childNode in xmlNode.ChildNodes)
            {
                ConvertXmlNodeToTreeNode(childNode, newTreeNode.Nodes);
            }
        }

    }
}