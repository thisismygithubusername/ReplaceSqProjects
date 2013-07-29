using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using FindandReplaceSql.Extensions;

namespace FindandReplaceSql.Models.ViewOutput
{
    public class SuspectLine
    {
        public SuspectLine()
        {
            Linetxt = new List<LineCharacter>();
        }
        public List<LineCharacter> Linetxt { get; set; }

        public void AddChar(char c, Color color)
        {
            Linetxt.Add(new LineCharacter(c, color));
        }

        public void AddChar(LineCharacter character)
        {
            Linetxt.Add(character);
        }

        public void Remove(int index)
        {
            Linetxt.RemoveAt(index);
        }

        public void Write(RichTextBox textbox)
        {
            foreach (var lineCharacter in Linetxt)
            {
                textbox.AppendText(lineCharacter.Value.ToString(CultureInfo.InvariantCulture), lineCharacter.Color);
            }
        }

        public override string ToString()
        {
            return Linetxt.Aggregate("", (current, lineCharacter) => current + lineCharacter.Value);
        }

    }
 
}
