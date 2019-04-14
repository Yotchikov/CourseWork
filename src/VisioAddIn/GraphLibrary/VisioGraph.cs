using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Graphviz4Net.Dot;
using Graphviz4Net.Dot.AntlrParser;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace GraphLibrary
{
    /// <summary>
    /// Класс, объединяющий библиотеку Graphviz4Net и объектную модель Visio
    /// </summary>
    public class VisioGraph
    {
        private GraphParser gParser = new GraphParser();
        private DotGraph<string> graph;
        private Dictionary<string, Visio.Shape> vertices = new Dictionary<string, Visio.Shape>();

        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="input">Код графа</param>
        public VisioGraph(string input)
        {
            graph = gParser.ParseGraphData(input);
        }

        /// <summary>
        /// Процедура отображения графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        public void PresentGraphInVisio(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);
            Visio.Document visioConnectors = visioDocs.OpenEx("Basic Flowchart Shapes (US units).vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);
            Dictionary<string, Visio.Master> visioMasters = getMasterShapes(visioStencil);

            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                var node = graph.AllVertices.ElementAt(i);

                string shape = graph.AllVertices.ElementAt(i).Attributes.ContainsKey("shape") ? node.Attributes["shape"] : "ELLIPSE";
                string label = graph.AllVertices.ElementAt(i).Attributes.ContainsKey("label") ? node.Attributes["label"] : node.Id;
                
                vertices.Add(node.Id, visioPage.Drop(visioMasters[shape.ToUpper()], 1+i/2.0, 11 - i/2.0));
                vertices[node.Id].Text = label;
                vertices[node.Id].Resize(Visio.VisResizeDirection.visResizeDirNW, -0.7, Visio.VisUnitCodes.visInches);
            }

            for (int i = 0; i < graph.VerticesEdges.Count(); ++i)
            {
                var edge = graph.VerticesEdges.ElementAt(i);
                Visio.Shape connector = visioPage.Drop(visioConnectors.Masters.get_ItemU("Dynamic connector"), 0, 0);
                connector.get_Cells("EndArrow").Formula = "=5";
                // connector.get_Cells()
                vertices[edge.Source.Id].AutoConnect(vertices[edge.Destination.Id], Visio.VisAutoConnectDir.visAutoConnectDirDown, connector);
            }
        }

        /// <summary>
        /// Процедура удаления графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница Visio</param>
        public void RemoveGraphInVisio(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                var node = graph.AllVertices.ElementAt(i);
                if (vertices.ContainsKey(node.Id))
                    vertices[node.Id].DeleteEx(1);
            }
        }

        private Dictionary<string, Visio.Master> getMasterShapes(Visio.Document visioStencil)
        {
            Dictionary<string, Visio.Master> result = new Dictionary<string, Visio.Master>();
            result.Add("TRIANGLE", visioStencil.Masters.get_ItemU(@"Triangle"));
            result.Add("SQUARE", visioStencil.Masters.get_ItemU(@"Square"));
            result.Add("PENTAGON", visioStencil.Masters.get_ItemU(@"Pentagon"));
            result.Add("HEXAGON", visioStencil.Masters.get_ItemU(@"Hexagon"));
            result.Add("OCTAGON", visioStencil.Masters.get_ItemU(@"Octagon"));
            result.Add("RECTANGLE", visioStencil.Masters.get_ItemU(@"Rectangle"));
            result.Add("RECT", visioStencil.Masters.get_ItemU(@"Rectangle"));
            result.Add("BOX", visioStencil.Masters.get_ItemU(@"Rectangle"));
            result.Add("CIRCLE", visioStencil.Masters.get_ItemU(@"Circle"));
            result.Add("ELLIPSE", visioStencil.Masters.get_ItemU(@"Ellipse"));
            result.Add("DIAMOND", visioStencil.Masters.get_ItemU(@"Diamond"));
            return result;
        }
    }
}
