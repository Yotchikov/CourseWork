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

            // Parsing graph code
            Parse(text);

        }

        private void Parse(string code)
        {
            // Reading file
            int i = 0;
            string stage = "Name";

            if (code[i] != ' ')
                throw new ArgumentException("Must be space between graph type and graph name");

            while (i != code.Length)
            {
                if (code[i] == ' ' || code[i] == '\n')
                {
                    ++i;
                    continue;
                }

                switch (stage)
                {
                    case "Name":
                        while (i != code.Length && code[i] != ' ' && code[i] != '{' && code[i] != '\n') ;
                        {
                            if (!Char.IsLetterOrDigit(code[i]))
                                throw new ArgumentException("Illegal name");
                            name += code[i++];
                        }
                        stage = "Open";
                        break;

                    case "Open":
                        if (code[i] != '{')
                            throw new ArgumentException("Missed be '{'");
                        stage = "Declaration";
                        ++i;
                        break;

                    case "Declaration":
                        string nodeName = String.Empty;
                        while (i != code.Length && code[i] != ' ' && code[i] != '-' && code[i] != ';' && code[i] != '\n' && code[i] != '[')
                        {
                            if (!Char.IsLetterOrDigit(code[i]))
                                throw new ArgumentException("Node name contains non alphanumeric character");
                            nodeName += code[i];
                            ++i;
                        }
                        switch (code[i])
                        {
                            case '-':
                                stage = "DeclarationOfEdge";
                                ++i;
                                break;
                            case ';':
                                nodes.Add(new Node(name));
                                ++i;
                                break;
                            case '\n':
                                stage = "StyleOfNode";
                                ++i;
                                break;
                        }

                        break;
                }
            }
        }

        public override string ToString()
        {
            return name + " " + (oriented ? "oriented" : "disoriented");
        }
    }
}
