using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using FindandReplaceSql.Models;

namespace FindandReplaceSql.Modules
{
    public class SuspectBlock
    {
        public SuspectBlock(AspPage page, int suspectLine)
        {
            LineNumber = suspectLine;
            SqlBlock = page.Lines.GetRange(LineNumber - 1, 10);
        }

        public int LineNumber { get; set; }

        public List<AspLine> SqlBlock { get; set; }

        public List<AspLine> GetBlockFromPage(AspPage page)
        {
            return page.Lines.Count >= LineNumber + 10
                       ? page.Lines.GetRange(LineNumber - 1, 10)
                       : page.Lines.GetRange(LineNumber - 1, page.Lines.Count - LineNumber);
        }

    }
}
