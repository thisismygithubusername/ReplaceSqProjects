using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models
{
    public class AspLine
    {
        public AspLine()
        {

        }
        public AspLine(int lineNum, string line)
        {
            LineNumber = lineNum;
            Line = line;
        }

        public int LineNumber { get; set; }

        public string Line { get; set; }

        public int StartDisplayLine
        {
            get
            {
                return (LineNumber > 5) ? LineNumber - 5 : 0;
            }
        }

        public int EndDisplayLine
        {
            get { return LineNumber + 10; }
        }

        public override string ToString()
        {
            return FormatLineNumer(LineNumber) + " : " + Line;
        }

        private string FormatLineNumer(int lineNumber)
        {
            if (lineNumber < 10)
            {
                return "   " + lineNumber;
            }
            if (lineNumber < 100)
            {
                return "  " + lineNumber;
            }
            if (lineNumber < 1000)
            {
                return " " + lineNumber;
            }
            return lineNumber + "";
        }
    }
}
