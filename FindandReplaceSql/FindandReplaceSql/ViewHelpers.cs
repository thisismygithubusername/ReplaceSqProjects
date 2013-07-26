using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindandReplaceSql
{
    public static class ViewHelpers
    {
        public static void FormatListBoxWithColor(this ListBox listBox, Color color)
        {
            listBox.ForeColor = Color.Red;
            listBox.HorizontalScrollbar = true;
            listBox.Items.Clear();
        }

        public static void SetTopIndexAndSelect(this ListBox listbox, int index)
        {
            listbox.TopIndex = index;
            listbox.SelectedIndex = index;
        }
    }
}
