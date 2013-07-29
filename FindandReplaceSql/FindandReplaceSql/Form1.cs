using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FindandReplaceSql.Models;
using FindandReplaceSql.Modules;
using FindandReplaceSql.Extensions;

namespace FindandReplaceSql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Button PrevButton
        {
            get { return this.button6; }
        }

        private Button NextButton
        {
            get { return this.button5; }
        }

        private Button WrapButton
        {
            get { return this.button3; }
        }

        private Button CustomButton
        {
            get { return this.button4; }
        }

        private RichTextBox RichDisplay
        {
            get { return this.richTextBox1; }
        }

        private string FileName
        {
            get; set;
        }

        private Stream FileStream
        {
            get; set;
        }

        private AspPage Page
        {
            get; set;
        }

        private int SuspectViewIndex
        {
            get; set;
        }

        private List<string> ToBeWrapped { get; set; }

        private int WrapIndex { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            var browser = new OpenFileDialog();
            browser.InitialDirectory = @"C:\IIS\wwwroot\mb-dev\Web\ASP";
            browser.ShowDialog();
            FileName = browser.FileName;
            DisplayTxtInBox(this.textBox1, browser.FileName);
        }

        //Search button 
        private void button2_Click(object sender, EventArgs e)
        {
            ClearAllAData();
            label2.ForeColor = Color.Blue;
            label2.Text = FileName.Split('\\').Last();
            Page = ASPParser.CreatePageFromFile(FileName);
            this.textBox4.Clear();
            this.textBox4.Text = Page.NumberofSuspects + "";
            DumpFileWithColors();
            PrintSuspectBlock(0);
        }

        private void ClearAllAData()
        {
            listBox2.Items.Clear();
            richTextBox1.Clear();
            Page = null;
            SuspectViewIndex = 0;

        }

        private void PrintSuspectBlock(int index)
        {
            if (Page.SuspectLines.Count > 0 && isValidIndex(index))
            {
                foreach (var line in new SuspectBlock(Page, Page.SuspectLines[index]).SqlBlock)
                {
                    listBox2.Items.Add(line);
                }
                AdjustDisplays(index);
            }
        }

        private bool isValidIndex(int index)
        {
            return index >= 0 && index < Page.SuspectLines.Count - 1;
        }

        private void AdjustDisplays(int index)
        {
            this.textBox3.Clear();
            this.textBox3.Text = (index + 1) + "";
            listBox2.SetTopIndexAndSelect(Page.SuspectLines[index] - 1);
            label6.Text = @"Conflict View: " + (index + 1);
            LoadRichTxt(Page.Lines[Page.SuspectLines[index] - 1].Line);
        }

        private void DisplayToBeWraped(string word)
        {
            this.richTextBox2.Clear();
            this.richTextBox2.Text = word;
        }

        private void LoadRichTxt(string line)
        {
            RichDisplay.Clear();
            var coloredLine = new LineAnalyzer { Line = line }.BuildColoredLine();
            ToBeWrapped = coloredLine.PossibleReplacements;

            foreach (var ch in coloredLine.Linetxt)
            {
                RichDisplay.AppendText(ch.Value.ToString(CultureInfo.InvariantCulture), ch.Color);
            }
            RichDisplay.Focus();
            DisplayToBeWraped(ToBeWrapped.Any() ? ToBeWrapped[0] : "");
        }

        private void DumpFileWithColors()
        {
            foreach (var aspline in Page.Lines)
            {
                this.listBox2.Items.Add(aspline);
            }
        }

        private void DisplayTxtInBox(TextBox box, string txt)
        {
            box.Clear();
            box.Text = txt;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.AutoSize = true;
            this.label4.Text = @"Conflict View:";
            this.textBox4.Clear();
            this.textBox4.Text = Page.SuspectLines + "";
        }

        //Next 
        private void button5_Click(object sender, EventArgs e)
        {
            if (SuspectViewIndex < Page.SuspectLines.Count)
            {
                SuspectViewIndex++;
                PrintSuspectBlock(SuspectViewIndex);
            }
        }

        //Previous
        private void button6_Click(object sender, EventArgs e)
        {
            if (SuspectViewIndex > 0)
            {
                SuspectViewIndex--;
                PrintSuspectBlock(SuspectViewIndex);
            }
        }

        //
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        //Todo 
        //CustomReplacetxtbox
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        //Wrap
        private void button3_Click(object sender, EventArgs e)
        {

        }

        //Suspect line 
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        //Suspeck Block display 
        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }
        //CustomReplace
        private void button4_Click(object sender, EventArgs e)
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


        private void label4_Click_1(object sender, EventArgs e)
        {
            var value = textBox3.Text;
            var intval = int.Parse(value);
            if(intval > 0 && intval <= Page.NumberofSuspects )
            {
                SuspectViewIndex = intval - 1;
                PrintSuspectBlock(SuspectViewIndex);
            }
            RichDisplay.Focus();           
        }


        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }


    }

    
}
