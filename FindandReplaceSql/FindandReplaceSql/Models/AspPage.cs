using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
