using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models
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

        public override string ToString()
        {
            return Linetxt.Aggregate("", (current, ch) => current + ch + "");
        }

    }
 
}
