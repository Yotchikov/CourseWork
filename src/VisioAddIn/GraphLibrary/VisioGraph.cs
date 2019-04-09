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
    public class VisioGraph
    {
        private GraphParser gParser = new GraphParser();
        private DotGraph<string> graph;
        
        private List<Visio.Shape> vertices = new List<Visio.Shape>();
        private List<Visio.Shape> edges = new List<Visio.Shape>();

        public VisioGraph(string input)
        {
            graph = gParser.ParseGraphData(input);
        }

        public void PresentGraphInVisio(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);
            Dictionary<string, Visio.Master> visioMasters = getMasterShapes(visioStencil);

            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                string shape = graph.AllVertices.ElementAt(i).Attributes["shape"];
                vertices.Add(visioPage.Drop(visioMasters[shape.ToUpper()], i, 11 - i));
                vertices[i].Text = graph.AllVertices.ElementAt(i).Id;
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
            result.Add("CIRCLE", visioStencil.Masters.get_ItemU(@"Circle"));
            result.Add("DIAMOND", visioStencil.Masters.get_ItemU(@"Diamond"));
            return result;
        }
    }
}
