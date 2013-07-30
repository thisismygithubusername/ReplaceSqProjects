using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindandReplaceSql.Models.ViewOutput
{
    public class Wrapper
    {
        public Wrapper(List<string> words)
        {
            Words = words;
            CurrentIndex = 0;
        }

        public List<string> Words { get; set; }

        public int CurrentIndex { get; set; }

        public string GetCurrent()
        { 
            return Words[CurrentIndex].Trim();
        }

        public bool Next()
        {
            if (CurrentIndex + 1 < Words.Count())
            {
                CurrentIndex++;
                return true;
            }
            return false;
        }

        public bool Prev()
        {
            if (CurrentIndex > 0)
            {
                CurrentIndex--;
                return true;
            }
            return false;
        }

        public int Count()
        {
            return Words.Count();
        }

        public bool Any()
        {
            return Words.Count >= 1;
        }

        public Change Wrap()
        {
            if (Words.Any())
            {
                var old = Words[CurrentIndex].Trim();
                return new Change(old, "sqlClean(" + old +")");
            }
            return null;
        }
    }
}
