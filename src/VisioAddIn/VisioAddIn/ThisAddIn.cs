using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using GraphLibrary;


namespace VisioAddIn
{
    public partial class ThisAddIn
    {
        private Dictionary<Visio.Page, VisioGraph> graphs = new Dictionary<Visio.Page, VisioGraph>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        
        public void ShowGraph(string input)
        {
            Visio.Documents visioDocs = Application.Documents;
            Visio.Page visioPage = Application.ActiveDocument.Pages.Add();

            graphs.Add(visioPage, new VisioGraph(input));
            graphs[visioPage].PresentGraphInVisio(visioDocs, visioPage);
        }

        public void RemovePageIfError()
        {
            Application.ActiveDocument.Pages[Application.ActiveDocument.Pages.Count].Delete(1);
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
