using System;
using System.Windows.Forms;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    public partial class DirectoryTree : Form
    {
        public DirectoryTree()
        {
            InitializeComponent();
        }

        private void DirectoryTree_Load(object sender, EventArgs e)
        {
            // Set the first node.
            TreeNode rootNode = new TreeNode(@"C:\");
            treeDirectory.Nodes.Add(rootNode);

            // Fill the first level and expand it.
            Fill(rootNode);
            treeDirectory.Nodes[0].Expand();
        }

        private void treeDirectory_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            // If a dummy node is found, remove it and read the 
            // real directory list.
            if (e.Node.Nodes[0].Text == "*")
            {
                e.Node.Nodes.Clear();
                Fill(e.Node);
            }
        }

        private void Fill(TreeNode dirNode)
        {
            DirectoryInfo dir = new DirectoryInfo(dirNode.FullPath);

            // An exception could be thrown in this code if you don't
            // have sufficient security permissions for a file or directory.
            // You can catch and then ignore this exception.
            foreach (DirectoryInfo dirItem in dir.GetDirectories())
            {
                // Add node for the directory.
                TreeNode newNode = new TreeNode(dirItem.Name);
                dirNode.Nodes.Add(newNode);
                newNode.Nodes.Add("*");
            }
        }
    }
}