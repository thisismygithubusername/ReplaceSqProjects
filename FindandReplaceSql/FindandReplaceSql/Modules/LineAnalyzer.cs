using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FindandReplaceSql.Models;
using MbUnit.Framework;


namespace FindandReplaceSql.Modules
{
    public class LineAnalyzer
    {
        public LineAnalyzer() { }

        public string Line 
        { 
            get; set;
        }

        public int CountQuotes()
        {
            return countChar('"', Line);
        }

        public int CountAnds()
        {
            return countChar('&', Line);
        }

        private int countChar(char c, string str)
        {
            return str.Count(f => f == c);
        }

        public SuspectLine BuildColoredLine()
        {
            return new ColoredStringBuilder(Line, CountQuotes(), CountAnds()).PaintFirstRun().Refine();
        }
    }
}
