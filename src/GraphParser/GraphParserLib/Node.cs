using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphParserLib
{
    public class Node
    {
        private string name;
        private string style;
        private List<Edge> edges;

        public Node(string name)
        {
            this.name = name;
        }

        public Node(string name, string attributes)
        {
            this.name = name;
            this.style = attributes;
        }

        public void AddNeighbour(Node second, string attributes)
        {
            edges.Add(new Edge(this, second, attributes));
        }
    }
}
