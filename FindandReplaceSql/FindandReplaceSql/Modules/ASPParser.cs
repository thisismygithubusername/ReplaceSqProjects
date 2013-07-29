using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using FindandReplaceSql.Models;

namespace FindandReplaceSql.Modules
{
    public class ASPParser
    {
        public ASPParser(string fileName)
        {
            FileName = fileName;
        }

        public string FileName
        {
            get; set;
        }

        public static AspPage CreatePageFromFile(string fileName)
        {  
            return new ASPParser(fileName).GetPageFromStream(new StreamReader(fileName));
        }

        public AspPage ParsePage()
        {
            return GetPageFromStream(new StreamReader(FileName));
        }

        private  AspPage GetPageFromStream(StreamReader fileStream )
        {
            string line = "";
            var pageLines = new List<AspLine>();
            var suspectLines = new List<int>();

            for (var count = 1; ((line = fileStream.ReadLine()) != null); count++)
            {
                if (CheckIfContainsSql(line)){ suspectLines.Add(count);}
                pageLines.Add(new AspLine { Line = line, LineNumber = count });
            }
            fileStream.Close();

            return new AspPage(pageLines, suspectLines);
        }

        private static bool CheckIfContainsSql(string line)
        {
            if (line.Contains("SELECT") || line.Contains("DELETE") ||
                line.Contains("FROM") || line.Contains("WHERE") ||
                line.Contains("strSQL") || line.Contains("SET"))
                return true;
            return false;
        }
        


        
    }
}
