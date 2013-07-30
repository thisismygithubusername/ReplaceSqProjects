using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FindandReplaceSql.Modules;

namespace FindandReplaceSql.Models
{
    public class AspPage
    {
        public AspPage(List<AspLine> lines, List<int> suspectedLines)
        {
            Lines = lines;
            SuspectLines = suspectedLines;
        }

        public List<AspLine> Lines { get; set; }

        public List<int> SuspectLines { get; set; } 

        public int NumberofSuspects
        {
            get
            {
                if (SuspectLines != null && SuspectLines.Count > 0)
                    return SuspectLines.Count;
                return 0;
            }
        }

        public void RefineSuspects()
        {
            var refinedSuspects = new List<int>();
            foreach (var suspectLineNum in SuspectLines)
            {
                if (new LineAnalyzer(Lines[suspectLineNum].Line).BuildColoredLine().PossibleReplacements.Any())
                {
                    refinedSuspects.Add(suspectLineNum);
                }
            }
            SuspectLines = refinedSuspects;
        }
    }
}
