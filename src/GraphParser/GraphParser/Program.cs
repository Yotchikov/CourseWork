using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GraphParserLib;

namespace GraphParser
{
    class Program
    {
        static void Main(string[] args)
        {
            Graph g = new Graph("digraph hello {");
            Console.WriteLine(g.ToString());
        }
    }
}
