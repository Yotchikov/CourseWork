using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using Graphviz4Net.Dot.AntlrParser;
using Graphviz4Net.Dot;

namespace VisioAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private DotGraph<string> ParseGraphData(string code)
        {
            var parser = AntlrParserAdapter<string>.GetParser();
            var result = parser.Parse(
                @"
                digraph {
                    node [label=""\N""];
                    graph [bb=""0,0,74,112""];
                    Hello [pos=""37,93"", width=""0.91667"", height=""0.52778""];
                    World [pos=""37,19"", width=""1.0278"", height=""0.52778""];
                    Hello -> World [pos=""e,37,38.249 37,73.943 37,66.149 37,56.954 37,48.338""];
                }");

            return result;
        }

        public void ShowGraph(string input)
        {

            Visio.Documents visioDocs = this.Application.Documents;
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);
            Visio.Page visioPage = this.Application.ActivePage;

            DotGraph<string> graph = ParseGraphData(input);

            Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Rectangle");
            List<Visio.Shape> vertices = new List<Visio.Shape>();
            
            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                vertices.Add(visioPage.Drop(visioRectMaster, i, 11-i));
                vertices[i].Text = graph.AllVertices.ElementAt(i).Id;
            }

            /*
            Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Rectangle");
            Visio.Shape visioRectShape = visioPage.Drop(visioRectMaster, 0, 0);
            visioRectShape.Text = @"Rectangle text.";

            Visio.Master visioHexagonMaster = visioStencil.Masters.get_ItemU(@"Hexagon");
            Visio.Shape visioHexagonShape = visioPage.Drop(visioHexagonMaster, 7.0, 5.5);
            visioHexagonShape.Text = @"Hexagon text.";*/
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
