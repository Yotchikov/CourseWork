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
        private bool isOriented;

        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="input">Код графа</param>
        public VisioGraph(string input)
        {
            graph = gParser.ParseGraphData(input);
            isOriented = input.StartsWith("digraph");
        }

        /// <summary>
        /// Процедура отображения графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        public void PresentGraphInVisio(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            // Название для страницы с графом
            if (graph.Attributes.ContainsKey("label"))
                visioPage.Name = graph.Attributes["label"];
            else
                visioPage.Name = "Graph";
            
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);
            Visio.Document visioConnectors = visioDocs.OpenEx("Basic Flowchart Shapes (US units).vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);
            
            // Мастер-объект базовых фигур Visio
            Dictionary<string, Visio.Master> visioMasters = GetMasterShapes(visioStencil);

            // Расстановка вершин графа
            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                var node = graph.AllVertices.ElementAt(i);

                string shape = node.Attributes.ContainsKey("shape") ? node.Attributes["shape"] : "ELLIPSE";
                string label = node.Attributes.ContainsKey("label") ? node.Attributes["label"] : node.Id;
                string color = node.Attributes.ContainsKey("color") ? node.Attributes["color"] : "black";
                string style = node.Attributes.ContainsKey("style") ? node.Attributes["style"] == "filled" ? "filled" : LineStyle(node.Attributes["style"].ToLower()) : "1";

                vertices.Add(node.Id, visioPage.Drop(visioMasters[shape.ToUpper()], 1+i/2.0, 11 - i/2.0));
                vertices[node.Id].Text = label;
                if (style == "filled")
                    vertices[node.Id].get_CellsSRC(
                (short)Visio.VisSectionIndices.visSectionObject,
                (short)Visio.VisRowIndices.visRowFill,
                (short)Visio.VisCellIndices.visFillForegnd).FormulaU = VisioColor.ColorToRgb(color.ToLower());
                else
                    vertices[node.Id].get_CellsU("LinePattern").FormulaU = style;
                vertices[node.Id].get_CellsU("LineColor").FormulaU = VisioColor.ColorToRgb(color.ToLower());
                vertices[node.Id].Resize(Visio.VisResizeDirection.visResizeDirNW, -0.8, Visio.VisUnitCodes.visInches);
            }

            // Соединение вершин графа ребрами
            for (int i = 0; i < graph.VerticesEdges.Count(); ++i)
            {
                var edge = graph.VerticesEdges.ElementAt(i);
                Visio.Shape connector = visioPage.Drop(visioConnectors.Masters.get_ItemU("Dynamic connector"), 0, 0);
                connector.get_Cells("ConLineRouteExt").FormulaU = "2";
                if (isOriented)
                    connector.get_Cells("EndArrow").Formula = "=5";
                else
                    connector.get_Cells("EndArrow").Formula = "=0";

                string label = edge.Attributes.ContainsKey("label") ? edge.Attributes["label"] : "";
                string color = edge.Attributes.ContainsKey("color") ? edge.Attributes["color"] : "black";
                string linestyle = edge.Attributes.ContainsKey("style") ? LineStyle(edge.Attributes["style"].ToLower()) : "1";

                connector.Text = label;
                connector.get_CellsU("LineColor").FormulaU = VisioColor.ColorToRgb(color.ToLower());
                connector.get_CellsU("LinePattern").FormulaU = linestyle;
                
                vertices[edge.Source.Id].AutoConnect(vertices[edge.Destination.Id], Visio.VisAutoConnectDir.visAutoConnectDirDown, connector);
            }

            masterConnector.Delete();
        }

        /// <summary>
        /// Процедура удаления графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница Visio</param>
        public void RemoveGraphInVisio(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            visioPage.Delete(1);
        }

        /// <summary>
        /// Процедура, сопоставляющая ключу-строке из допустимых фигур DOT мастер-фигуру Visio
        /// </summary>
        /// <param name="visioStencil"></param>
        /// <returns>Словарь string-Visio.Master</returns>
        private Dictionary<string, Visio.Master> GetMasterShapes(Visio.Document visioStencil)
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
            result.Add("OVAL", visioStencil.Masters.get_ItemU(@"Ellipse"));
            result.Add("DIAMOND", visioStencil.Masters.get_ItemU(@"Diamond"));
            return result;
        }

        /// <summary>
        /// Процедура, возвращающая номер в формуле стиля линии
        /// </summary>
        /// <param name="styleString">Стиль в строковом формате</param>
        /// <returns>Номер формулы</returns>
        private string LineStyle(string styleString)
        {
            switch (styleString)
            {
                case "dashed":
                    return "2";
                case "dotted":
                    return "3";
                default:
                    return "1";
            }
        }
    }
}
