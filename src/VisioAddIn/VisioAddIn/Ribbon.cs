using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            Globals.ThisAddIn.ShowGraph("");
        }
    }
}
