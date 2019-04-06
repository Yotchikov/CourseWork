using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace VisioAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.Documents.Add("");

            Visio.Documents visioDocs = this.Application.Documents;
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);

            Visio.Page visioPage = this.Application.ActivePage;

            Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Rectangle");
            Visio.Shape visioRectShape = visioPage.Drop(visioRectMaster, 4.25, 5.5);
            visioRectShape.Text = @"Rectangle text.";

            Visio.Master visioHexagonMaster = visioStencil.Masters.get_ItemU(@"Hexagon");
            Visio.Shape visioHexagonShape = visioPage.Drop(visioHexagonMaster, 7.0, 5.5);
            visioHexagonShape.Text = @"Hexagon text.";
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
