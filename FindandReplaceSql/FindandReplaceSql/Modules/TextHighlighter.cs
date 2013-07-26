using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindandReplaceSql.Modules
{
    public class TextHighlighter
    {
        public void doSHit(RichTextBox textBox)
        {
            //textBox.SelectionStart = first;
            //textBox.SelectionLength = length;
            //scroll to the caret
            //textBox.ScrollToCaret();            
        }

        public static void TextHighlight(RichTextBox box, string line)
        {
            var coloredLine = new LineAnalyzer {Line = line}.BuildColoredLine().Linetxt;
            foreach (var ch in coloredLine)
            {
                box.SelectionStart = box.TextLength;
                box.SelectionLength = 0;

                box.SelectionColor = ch.Color;
                box.AppendText(ch.Value.ToString(CultureInfo.InvariantCulture));
                box.SelectionColor = box.ForeColor;
            }
        }
        

    }
}
