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

        private int NumBreaks
        {
            get { return CountAnds(); }
        }

        public SuspectLine BuildColoredLine()
        {
            var line = new SuspectLine();
           // this.Line =
             //   "strSQL = \"SELECT tblClassSch.ClassID, tblClassSch.ClassDate, tblClassSch.TrainerID, tblClassSch.PayScaleID, tblClassSch.StartTime, tblClassSch.EndTime, tblClasses.LocationID, tblClassDescriptions.VisitTypeID, tblClassDescriptions.ClassPayment, tblVisitTypes.NumDeducted FROM tblClasses INNER JOIN tblClassSch ON tblClasses.ClassID = tblClassSch.ClassID INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN tblVisitTypes ON tblClassDescriptions.VisitTypeID = tblVisitTypes.TypeID \"";
            bool printingString = false; 

            foreach (var letter in Line)
            {
                if (letter.Equals('"') )
                {
                    if (printingString == false)
                    {
                        printingString = true;
                        line.AddChar(CreateChar(true, letter));
                    }
                    else
                    {
                        line.AddChar(CreateChar(true, letter));
                        printingString = false;
                    }
                }
                else
                {
                    line.AddChar(CreateChar(printingString, letter));
                }
            }
            return line;
        }

        private LineCharacter CreateChar(bool printingState, char letter)
        {
            if (printingState)
            {
                return new LineCharacter(letter, Color.Coral);
            }
            else
            {
                return new LineCharacter(letter, Color.Black);
            }
        }

        [SetUp]
        public void SetUp()
        {
            
        }

        [TearDown]
        public void TearDown()
        {
            
        }


        // if two in between, if one after. 
        //int count = source.Count(f => f == '/');
        //int count = source.Length - source.Replace("/", "").Length;
        //int count = source.Split('/').Length - 1;
    }
}
