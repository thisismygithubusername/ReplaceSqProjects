using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models
{
    public class ModifiedLine : AspLine 
    {
        public override string ToString()
        {
            return this.Line;
        }

    }
}
