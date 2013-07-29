using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindandReplaceSql.Extensions
{
    public static class ElementExtensions
    {
        public static bool IsQuotes(this char ch)
        {
            return ch.Equals('"');
        }

        public static string RemoveWhiteSpace(this string str)
        {
            return str.Replace(" ", "");
        }

        public static void AppendText(this RichTextBox box, string str, Color color)
        {

                box.SelectionStart = box.TextLength;
                box.SelectionLength = 0;
                box.SelectionColor = color;
                box.AppendText(str);
                box.SelectionColor = box.ForeColor;
            
        }
    }
}
