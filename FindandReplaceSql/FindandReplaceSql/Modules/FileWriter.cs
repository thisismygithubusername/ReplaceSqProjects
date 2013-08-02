using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FindandReplaceSql.Models;

namespace FindandReplaceSql.Modules
{
    public class FileWriter
    {
        public FileWriter(ProgramSession session, string file)
        {
            Page = session.Page;
            LinesToReplace = session.ChangedLines;
            File = file;
        }

        public AspPage Page { get; set; }

        public string File { get; set; }

        public List<ModifiedLine> LinesToReplace { get; set; }

        private string SavePath
        {
            get { return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); }
        }

        private string BuildSavePath()
        {
            var path = Path.Combine(SavePath, "SqlCleanedFiles");
            if (Directory.Exists(path))
            {
                return path;
            }
            Directory.CreateDirectory(path);
            return path;
        }

        public void WriteChangesToFile()
        {
            var newPath = Path.Combine(BuildSavePath(),File.Split('\\').Last());
            using (var writer = new StreamWriter(newPath))
            {
                var replaceLines = LinesToReplace.GetEnumerator();
                replaceLines.MoveNext();
                for (int currentLine = 1; currentLine <= Page.Lines.Count; ++currentLine)
                {
                    if (replaceLines.Current!= null && currentLine == replaceLines.Current.LineNumber)
                    {
                        writer.WriteLine(replaceLines.Current.ChangedLine);
                        replaceLines.MoveNext();
                    }
                    else
                    {
                        writer.WriteLine(Page.Lines[currentLine - 1].Line);
                    }
                }
            }
        }
    }
}
