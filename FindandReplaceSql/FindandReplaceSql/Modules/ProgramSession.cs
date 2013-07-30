using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FindandReplaceSql.Models;
using FindandReplaceSql.Models.ViewOutput;

namespace FindandReplaceSql.Modules
{
    public class ProgramSession
    {
        public ProgramSession()
        {
            ChangedLines = new List<ModifiedLine>();
        }

        public List<ModifiedLine> ChangedLines { get; set; }

        public Wrapper Wrapper { get; set; }

        public AspPage Page { get; set; }

        public void ParsePage(string fileName)
        {
            Page = ASPParser.CreatePageFromFile(fileName);

        }

        public int SuspectViewIndex
        {
            get;
            set;
        }

        public ModifiedLine CurrentModifiedLine { get; set; }


        public int TotolReplacements { get; set; }
       
        public int WrapAndSave()
        {
            var wrapped = Wrapper.Wrap();

            if (wrapped == null)
            {
                return 0;
            }
            if (Wrapper.CurrentIndex == 0)
            {
                var pagelineindex = Page.SuspectLines[SuspectViewIndex];
                var oldLine = Page.Lines[pagelineindex];
                CurrentModifiedLine = new ModifiedLine(oldLine.Line, oldLine.LineNumber);
            }

            CurrentModifiedLine.ChangedText.Add(wrapped);
            if (Wrapper.CurrentIndex + 1 == Wrapper.Count())
            {
                ChangedLines.Add(CurrentModifiedLine);
                return 2;
            }
            return 1;

        }
    }
}
