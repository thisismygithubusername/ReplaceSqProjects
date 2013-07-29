using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FindandReplaceSql.Models.ViewOutput
{
    public class AnalyzedLine : SuspectLine
    {
        public AnalyzedLine(List<LineCharacter> coloredLine)
        {
            PossibleReplacements = new List<string>();
            ReplacedStrings = new List<string>();
            this.Linetxt = coloredLine;
        }

        public AnalyzedLine(List<LineCharacter> coloredLine, List<string> reds )
        {
            PossibleReplacements = reds;
            ReplacedStrings = new List<string>();
            this.Linetxt = coloredLine;
        }

        public List<string> PossibleReplacements { get; set; }

        public List<string> ReplacedStrings { get; set; }

        private MatchCollection GetAmpIndexes()
        {
            var amps = new Regex("&");
            return amps.Matches(this.ToString());
        }

    }
}
