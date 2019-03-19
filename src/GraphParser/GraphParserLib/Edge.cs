using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphParserLib
{
    public class Edge
    {
        private Node first;
        private Node second;
        private string style;

        public Edge(Node first, Node second, string attributes)
        {
            this.first = first;
            this.second = second;
            this.style = attributes;
        }
    }
}
