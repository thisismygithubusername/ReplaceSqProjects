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
            var path = Path.Combine(SavePath, "SqlCleaneFiles");
            if (Directory.Exists(path))
            {
                return path;
            }
            Directory.CreateDirectory(path);
            return path;
        }

        public void WriteChangesToFile()
        {
            var oldPath = File;
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

        public void CostlyReplace()
        {
            int line_to_edit = 2; // Warning: 1-based indexing!
            string sourceFile = "source.txt";
            string destinationFile = "target.txt";

            // Read the appropriate line from the file.
            string lineToWrite = null;
            using (StreamReader reader = new StreamReader(sourceFile))
            {
                for (int i = 1; i <= line_to_edit; ++i)
                    lineToWrite = reader.ReadLine();
            }

            if (lineToWrite == null)
                throw new InvalidDataException("Line does not exist in " + sourceFile);

            // Read the old file.
            string[] lines = null;// File.ReadAllLines(destinationFile);

            // Write the new file over the old file.
            using (StreamWriter writer = new StreamWriter(destinationFile))
            {
                for (int currentLine = 1; currentLine <= lines.Length; ++currentLine)
                {
                    if (currentLine == line_to_edit)
                    {
                        writer.WriteLine(lineToWrite);
                    }
                    else
                    {
                        writer.WriteLine(lines[currentLine - 1]);
                    }
                }
            }
        }
        public void CreateNewFile()
        {
            int line_to_edit = 2;
            string sourceFile = "source.txt";
            string destinationFile = "target.txt";
            string tempFile = "target2.txt";

            // Read the appropriate line from the file.
            string lineToWrite = null;
            using (StreamReader reader = new StreamReader(sourceFile))
            {
                for (int i = 1; i <= line_to_edit; ++i)
                    lineToWrite = reader.ReadLine();
            }

            if (lineToWrite == null)
                throw new InvalidDataException("Line does not exist in " + sourceFile);

            // Read from the target file and write to a new file.
            int line_number = 1;
            string line = null;
            using (StreamReader reader = new StreamReader(destinationFile))
            using (StreamWriter writer = new StreamWriter(tempFile))
            {
                while ((line = reader.ReadLine()) != null)
                {
                    if (line_number == line_to_edit)
                    {
                        writer.WriteLine(lineToWrite);
                    }
                    else
                    {
                        writer.WriteLine(line);
                    }
                    line_number++;
                }
            }

            // TODO: Delete the old file and replace it with the new file here.
        }
    }
}
