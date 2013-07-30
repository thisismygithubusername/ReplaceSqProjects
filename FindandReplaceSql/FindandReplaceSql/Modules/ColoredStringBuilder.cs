using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FindandReplaceSql.Models;
using FindandReplaceSql.Extensions;
using FindandReplaceSql.Models.ViewOutput;

namespace FindandReplaceSql.Modules
{
    public class ColoredStringBuilder
    {
        private string Txt { get; set; }

        public ColoredStringBuilder(string txt)
        {
            Txt = txt;
        }

        public PaintedString PaintFirstRun()
        {
            return new PaintedString(ColorString());
        }

        private SuspectLine ColorString()
        {
            var line = new SuspectLine();
            int quoteOccur = 0;

            foreach (var letter in Txt)
            {
                if (!letter.IsQuotes())
                {
                    line.AddChar(TxtGen(letter, (quoteOccur%2 == 1)));
                }
                else
                {
                    quoteOccur++;
                    line.AddChar(QuoteGen());
                }
            }
            return line;
        }

        private LineCharacter QuoteGen()
        {
            return new LineCharacter('"', Color.SaddleBrown);
        }
         
        private LineCharacter TxtGen(char symbol, bool strtype)
        {
            return strtype ? new LineCharacter(symbol, Color.Coral) : new LineCharacter(symbol, Color.Black);
        }

        public class PaintedString 
        {
            internal PaintedString(SuspectLine line)
            {
                Line = line;
            }

            public SuspectLine Line { get; set; }

            public AnalyzedLine Refine()
            {
                var linepieces = Line.ToString().Split('&');
                var colorsLists = new List<Color>();
                var plists = new List<string>();
                if (linepieces.Count() > 1)
                {
                    for (int i = 0; i < linepieces.Count(); i++)
                    {
                        var current = linepieces[i].Trim();

                        if (!IsProhibitedtoClean(current))
                        {
                            var isCleaned = IsCleaned(current);
                            colorsLists.Add(isCleaned ? Color.Green : Color.Red);
                            if (!isCleaned)
                            {
                                plists.Add(linepieces[i]);
                            }
                        }
                        else
                        {
                            colorsLists.Add(Color.White);
                        }
                    }
                }

                return ReColor(colorsLists, plists);
            }

            public bool IsCleaned(string chunk)
            {
                return chunk.Contains("sqlInject") || chunk.Contains("sqlClean(");
            }

            private bool IsProhibitedtoClean(string chunk)
            {
                if (IsChunkStringType(chunk))
                    return true;
                if (ContainsSqlandStr(chunk))
                    return true;
                if (ContainsWordsToAvoid(chunk))
                    return true;
                if (IsUnderscore(chunk))
                    return true;
                if (ContainsOnlyOneQuote(chunk))
                    return true;
                return false;
            }

            private bool ContainsOnlyOneQuote(string chunk)
            {
                return countChar('"', chunk) == 1;
            }

            private bool IsChunkStringType(string chunk)
            {
                return chunk.First().Equals('"') && chunk.Last().Equals('"');
            }

            private bool ContainsSqlandStr(string chunk)
            {
                return chunk.Contains("SQL") || chunk.Contains("str");
            }

            private bool IsUnderscore(string chunk)
            {
                return chunk.Trim().Equals("_");
            }

            private int countChar(char c, string str)
            {
                return str.Count(f => f == c);
            }

            private bool ContainsWordsToAvoid(string chunk)
            {
                foreach (var word in WordsToAvoid)
                {
                    if (chunk.Contains(word))
                    {
                        return true;
                    }
                }
                return false;
            }

            private AnalyzedLine ReColor(List<Color> colorList, List<string> replacements)
            {
                int currentwordindex = 0;
                if (colorList.Any())
                {
                    foreach (LineCharacter t in Line.Linetxt)
                    {
                        if (t.Value.Equals('&'))
                        {
                            currentwordindex++;
                        }
                        else if (!colorList[currentwordindex].Equals(Color.White))
                        {
                            t.Color = colorList[currentwordindex];
                        }
                    }
                }
                return new AnalyzedLine(Line.Linetxt, replacements);
            }

            private readonly string[] WordsToAvoid =  {"step", "Step", "Sep"};
        }
    }

}
