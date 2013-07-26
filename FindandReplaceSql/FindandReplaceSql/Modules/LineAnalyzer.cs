using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Modules
{
    public class LineAnalyzer
    {
        public LineAnalyzer() { }

        public string Line { get; set;
        }

        public void CountQuotes()
        {
            countChar('"', Line);
        }

        public int CountAnds()
        {
            return countChar('/', Line);
        }

        private int countChar(char c, string str)
        {
            return str.Count(f => f == c);
        }


        // if two in between, if one after. 
        //int count = source.Count(f => f == '/');
        //int count = source.Length - source.Replace("/", "").Length;
        //int count = source.Split('/').Length - 1;
    }
}
