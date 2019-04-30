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
        
        /// <summary>
        /// Метод отображения графа в Visio
        /// </summary>
        /// <param name="input"></param>
        public void ShowGraph(string input)
        {
            Application.ActiveDocument.Pages.BeforePageDelete += Globals.ThisAddIn.DeleteGraph;

            Visio.Documents visioDocs = Application.Documents;
            Visio.Page visioPage = Application.ActiveDocument.Pages.Add();

            graphs.Add(visioPage, new VisioGraph(input));
            graphs[visioPage].PresentGraphInVisio(visioDocs, visioPage);

            Application.BeforeShapeDelete += DeleteShape;
            Application.ActivePage.ConnectionsAdded += AddEdge;
        }

        /// <summary>
        /// Метод удаления фигуры
        /// </summary>
        /// <param name="Shape"></param>
        private void DeleteShape(Visio.Shape Shape)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Метод удаления ребра
        /// </summary>
        /// <param name="Connects"></param>
        private void AddEdge(Visio.Connects Connects)
        {
            if (graphs.ContainsKey(Application.ActivePage))
                graphs[Application.ActivePage].AddEdge(Connects);
        }

        /// <summary>
        /// Метод удаления страницы, в случае, если возникла ошибка
        /// </summary>
        public void RemovePageIfError()
        {
            Application.ActiveDocument.Pages[Application.ActiveDocument.Pages.Count].Delete(1);
        }

        /// <summary>
        /// Удаление графа из словаря, если была удалена страница
        /// </summary>
        /// <param name="Page"></param>
        private void DeleteGraph(Visio.Page Page)
        {
            if (graphs.ContainsKey(Page))
            {
                graphs.Remove(Page);
            }
        }

        /// <summary>
        /// Метод экспорта графа в файл
        /// </summary>
        /// <param name="filePath"></param>
        public void ExportGraph(string filePath)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                graphs[Application.ActivePage].ExportGraph(filePath);
            }
            else throw new ArgumentException("На данной странице не представлен граф");
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
