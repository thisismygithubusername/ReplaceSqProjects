﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models
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
