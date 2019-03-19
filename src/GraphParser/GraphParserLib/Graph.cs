using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphParserLib
{
    public class Graph
    {
        private bool oriented;
        private string name;

        private List<Node> nodes;

        public Graph(string text)
        {
            // Checking orientation of the graph
            if (!text.ToLower().StartsWith("graph") && !text.ToLower().StartsWith("digraph"))
                throw new ArgumentException("File must start with 'graph' or 'digraph'");
            if (text.ToLower()[0] == 'g')
                oriented = false;
            else
                oriented = true;

            // Deleting type of graph
            text = text.Substring(oriented ? 7 : 5);

            // Reading file
            int i = 0;
            string stage = "Name";

            if (text[i] != ' ')
                throw new ArgumentException("Must be space between graph type and graph name");
            while (i != text.Length)
            {
                if (text[i] == ' ')
                {
                    ++i;
                    continue;
                }

                if (stage == "Name")
                {
                    while (i != text.Length && text[i] != ' ')
                    {
                        if (!Char.IsLetterOrDigit(text[i]))
                            throw new ArgumentException("Illegal name!");
                        name += text[i++];
                    }
                    stage = "Open";
                }
            }

        }

        public override string ToString()
        {
            return name + " " + (oriented ? "oriented" : "disoriented");
        }
    }
}
