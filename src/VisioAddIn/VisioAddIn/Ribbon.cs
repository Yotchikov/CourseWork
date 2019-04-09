using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace VisioAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void openFileButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string input = File.ReadAllText(openFileDialog.FileName);
                Globals.ThisAddIn.ShowGraph(input);
            }
        }
    }
}
