using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindandReplaceSql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var browser = new OpenFileDialog();
            browser.InitialDirectory = @"C:\IIS\wwwroot\mb-dev";

            var file = browser.ShowDialog();
            DisplayTxtInBox(this.textBox1, browser.FileName);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        /************************* helpers *****************************/
        private void DisplayTxtInBox(TextBox box, string txt)
        {
            box.Clear();
            box.Text = txt;
        }
    }
}
