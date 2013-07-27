using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
