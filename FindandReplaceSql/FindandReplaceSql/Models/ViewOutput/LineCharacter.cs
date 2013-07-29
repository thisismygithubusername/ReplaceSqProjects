using System.Drawing;

namespace FindandReplaceSql.Models.ViewOutput
{
    public class LineCharacter
    {
        public LineCharacter(char c, Color color)
        {
            Value = c;
            Color = color;
        }

        public char Value { get; set; }
        public Color Color { get; set; }
    }
}
