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
using FindandReplaceSql.Models;
using FindandReplaceSql.Modules;

namespace FindandReplaceSql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string FileName { get; set; }
        private Stream FileStream { get; set; }
        private AspPage Page { get; set; }
        private int SuspectViewIndex { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            var browser = new OpenFileDialog();
            browser.InitialDirectory = @"C:\IIS\wwwroot\mb-dev";

            browser.ShowDialog();
            FileName = browser.FileName;
            DisplayTxtInBox(this.textBox1, browser.FileName);
        }


        /************************* helpers *****************************/

        private void DisplayTxtInBox(TextBox box, string txt)
        {
            box.Clear();
            box.Text = txt;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //Search button 
        private void button2_Click(object sender, EventArgs e)
        {
            label2.ForeColor = Color.Blue;
            label2.Text = FileName.Split('\\').Last();
            //  + ": " + FileName;
            this.Page = ASPParser.ExamFile(FileName);
            DumpFileWithColors();
            PrintSuspectBlock(0);
        }

        private void PrintSuspectBlock(int index)
        {
            this.listBox1.ForeColor = Color.Red;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.Items.Clear();

            if (Page.SuspectLines.Count > 0)
            {
                foreach (var line in new SuspectBlock(Page, Page.SuspectLines[index]).SqlBlock)
                {
                    this.listBox1.Items.Add(line);
                }
                this.listBox2.TopIndex = Page.SuspectLines[index] - 1;
                this.listBox2.SelectedIndex = Page.SuspectLines[index] - 1;
                this.label6.Text = "Conflict View: " + (index + 1 );
            }
        }

        private void DumpFileWithColors()
        {
            foreach (var aspline in Page.Lines)
            {
                this.listBox2.Items.Add(aspline);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.AutoSize = true;
            label6.Text = "Conflict View: " + SuspectViewIndex + 1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SuspectViewIndex++;
            PrintSuspectBlock(SuspectViewIndex);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

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


    }

    
}
