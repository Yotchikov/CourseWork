using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Graphviz4Net.Dot;
using Graphviz4Net.Dot.AntlrParser;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.IO;

namespace GraphLibrary
{
    /// <summary>
    /// Класс, объединяющий библиотеку Graphviz4Net и объектную модель Visio
    /// </summary>
    public class VisioGraph
    {
        private GraphParser gParser = new GraphParser();
        private DotGraph<string> graph;
        private Dictionary<DotVertex<string>, Visio.Shape> vertices = new Dictionary<DotVertex<string>, Visio.Shape>();
        private bool isOriented;
        private Dictionary<Visio.Shape, List<Visio.Shape>> newEdges = new Dictionary<Visio.Shape, List<Visio.Shape>>();

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
            PresentNodes(visioDocs, visioPage);
            PresentEdges(visioDocs, visioPage);

            visioPage.Layout();
        }

        /// <summary>
        /// Процедура отображения вершин графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        private void PresentNodes(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);

            // Мастер-объект базовых фигур Visio
            Dictionary<string, Visio.Master> visioMasters = GetMasterShapes(visioStencil);

            // Расстановка вершин графа
            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                // Вершина
                var node = graph.AllVertices.ElementAt(i);

                // Стили вершины
                string shape = node.Attributes.ContainsKey("shape") ? node.Attributes["shape"] : "ELLIPSE";
                string label = node.Attributes.ContainsKey("label") ? node.Attributes["label"] : node.Id;
                string color = node.Attributes.ContainsKey("color") ? node.Attributes["color"] : "black";
                string style = node.Attributes.ContainsKey("style") ? node.Attributes["style"] == "filled" ? "filled" : LineStyle(node.Attributes["style"].ToLower()) : "1";

                // Добавление вершины на страницу Visio
                vertices.Add(node, visioPage.Drop(visioMasters[shape.ToUpper()], 1 + i / 2.0, 11 - i / 2.0));

                // Установка стилей для фигуры на странице Visio
                vertices[node].Text = label;
                if (style == "filled")
                    vertices[node].get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillForegnd).FormulaU = VisioColor.ColorToRgb(color.ToLower());
                else
                    vertices[node].get_CellsU("LinePattern").FormulaU = style;
                vertices[node].get_CellsU("LineColor").FormulaU = VisioColor.ColorToRgb(color.ToLower());

                // Ресайзинг
                vertices[node].Resize(Visio.VisResizeDirection.visResizeDirNW, -0.8, Visio.VisUnitCodes.visInches);
            }
        }

        /// <summary>
        /// Процедура отображения ребер графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        private void PresentEdges(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            Visio.Document visioConnectors = visioDocs.OpenEx("Basic Flowchart Shapes (US units).vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);

            // Соединение вершин графа ребрами
            for (int i = 0; i < graph.VerticesEdges.Count(); ++i)
            {
                // Ребро
                var edge = graph.VerticesEdges.ElementAt(i);

                // Фигура соединидельной линии (коннектора)
                Visio.Shape connector = visioPage.Drop(visioConnectors.Masters.get_ItemU("Dynamic connector"), 0, 0);
                connector.get_Cells("ConLineRouteExt").FormulaU = "2";
                if (isOriented)
                    connector.get_Cells("EndArrow").Formula = "=5";
                else
                    connector.get_Cells("EndArrow").Formula = "=0";

                // Стили ребра
                string label = edge.Attributes.ContainsKey("label") ? edge.Attributes["label"] : "";
                string color = edge.Attributes.ContainsKey("color") ? edge.Attributes["color"] : "black";
                string linestyle = edge.Attributes.ContainsKey("style") ? LineStyle(edge.Attributes["style"].ToLower()) : "1";

                // Установка стилей для фигуры на странице Visio
                connector.Text = label;
                connector.get_CellsU("LineColor").FormulaU = VisioColor.ColorToRgb(color.ToLower());
                connector.get_CellsU("LinePattern").FormulaU = linestyle;

                // Соединение вершин при помощи данного коннектора
                vertices[edge.Source].AutoConnect(vertices[edge.Destination], Visio.VisAutoConnectDir.visAutoConnectDirDown, connector);

                // Удаление коннектора-болванки
                connector.Delete();
            }
        }

        public void ChangeLabel(Visio.Shape shape)
        {
            if (vertices.ContainsValue(shape))
                foreach (var node in vertices)
                    if (node.Value == shape)
                    {
                        foreach (var v in graph.AllVertices)
                            if (v == node.Key)
                            {
                                if (v.Attributes.ContainsKey("label"))
                                    v.Attributes["label"] = shape.Text;
                                else
                                    v.Attributes.Add("label", shape.Text);
                                break;
                            }
                        break;
                    }
        }

        /// <summary>
        /// Метод, вызывающийся при добавлении нового ребра между ребрами графа
        /// </summary>
        /// <param name="connects">Коннектор</param>
        public void AddEdge(Visio.Connects connects)
        {
            // Если соединение добавлено с использованием существующей соединительной линии, т.е. добавляем вторую вершину
            if (newEdges.ContainsKey(connects.FromSheet))
            {
                // Добавляем вторую вершину в пару связанных вершин
                newEdges[connects.FromSheet].Add(connects.ToSheet);

                // Если обе соединенные фигуры - это вершины нашего графа
                if (vertices.ContainsValue(newEdges[connects.FromSheet][0]) && vertices.ContainsValue(newEdges[connects.FromSheet][1]))
                {
                    DotVertex<string> from = new DotVertex<string>("");
                    DotVertex<string> to = new DotVertex<string>("");

                    // Ищем, каким вершинам они соответствуют
                    foreach (var node in vertices)
                    {
                        if (node.Value == newEdges[connects.FromSheet][0])
                        {
                            from = node.Key;
                        }
                        if (node.Value == newEdges[connects.FromSheet][1])
                        {
                            to = node.Key;
                        }
                    }

                    // Добавляем ребро с данными вершинами
                    graph.AddEdge(new DotEdge<string>(from, to));
                }

                // Удаляем пару
                newEdges.Remove(connects.FromSheet);
            }
            else
            {
                // Применяем стили к коннектору
                connects.FromSheet.get_Cells("ConLineRouteExt").FormulaU = "2";
                connects.FromSheet.get_Cells("EndArrow").Formula = "=5";

                // Добавляем новый коннектор в словарь + вершину, от которой он стартовал, ожидая соединения ее со второй вершиной
                List<Visio.Shape> listOfConnectingNodes = new List<Visio.Shape>();
                listOfConnectingNodes.Add(connects.ToSheet);
                newEdges.Add(connects.FromSheet, listOfConnectingNodes);
            }
        }

        /// <summary>
        /// Метод, вызывающийся при удалении ребра
        /// </summary>
        /// <param name="connects"></param>
        public void DeleteEdge(Visio.Connects connects)
        {
            throw new NotImplementedException();
            if (vertices.ContainsValue(connects.FromSheet) && vertices.ContainsValue(connects.ToSheet))
            {
                for (int i = 0; i < graph.VerticesEdges.Count(); ++i)
                {
                    // Ребро
                    var edge = graph.VerticesEdges.ElementAt(i);

                    if (vertices[edge.Source] == connects.FromSheet && vertices[edge.Destination] == connects.ToSheet)
                    {
                        graph.RemoveEdge(edge);
                        break;
                    }
                }
            }
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

        /// <summary>
        /// Процедура экспорта графа в файл
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        public void ExportGraph(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                new GraphToDotConverter().Convert(sw, graph, new AttributesProvider());
            }
        }
    }

    public class AttributesProvider : IAttributesProvider
    {
        public IDictionary<string, string> GetVertexAttributes(object vertex)
        {
            return new Dictionary<string, string>();
        }
    }
}
