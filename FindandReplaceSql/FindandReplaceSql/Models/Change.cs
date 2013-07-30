using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models
{
    public class Change
    {
        public Change(string old, string replaced )
        {
            Old = old;
            Replaced = replaced;
        }

        public string Old { get; set; }

        public string Replaced { get; set; }
    }
}
