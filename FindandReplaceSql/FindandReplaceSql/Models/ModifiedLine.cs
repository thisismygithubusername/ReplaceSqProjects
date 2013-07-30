using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models
{
    public class ModifiedLine : AspLine 
    {
        public ModifiedLine(string line, int lineNum)
        {
            Line = line;
            LineNumber = lineNum;
            ChangedText = new List<Change>();
        }

        public override string ToString()
        {
            return FormatLineNumer(LineNumber) + " : " + BuildChanged();
        }

        public string ChangedLine { get; set; }

        public List<Change> ChangedText { get; set; }

        private string BuildChanged()
        {
            ChangedLine = Line;

            foreach (var change in ChangedText)
            {
                ChangedLine = ChangedLine.Replace(change.Old, change.Replaced);
            }
            return ChangedLine;
        }

    }
}
