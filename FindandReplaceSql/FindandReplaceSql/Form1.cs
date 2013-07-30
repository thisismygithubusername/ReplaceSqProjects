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
using FindandReplaceSql.Models.ViewOutput;
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

        private TextBox TbFileName
        {
            get { return this.textBox1; }
        }

        private TextBox TbConflictNumber
        {
            get { return this.textBox3; }
        }

        private TextBox TbTotalConflictNum
        {
            get { return this.textBox4; }
        }

        private ListBox LbFileView
        {
            get { return this.listBox2; }
        }
        private ListBox LbChangedView
        {
            get { return this.listBox1; }
        }

        private RichTextBox RichDisplayLine
        {
            get { return this.richTextBox1; }
        }

        private RichTextBox RichDisplayWrap
        {
            get { return this.richTextBox2; }
        }

        private string FileName
        {
            get; set;
        }

        private Label CurrentReplacement
        {
            get { return this.label8; }
        }

        private Label TotalReplacements
        {
            get { return this.label10; }
        }


        private ProgramSession Session { get; set; }

        //Select Button 
        private void button1_Click(object sender, EventArgs e)
        {
            var browser = new OpenFileDialog {InitialDirectory = @"C:\IIS\wwwroot\mb-dev\Web\ASP"};
            browser.ShowDialog();
            FileName = browser.FileName;
            DisplayTxtInBox(TbFileName, FileName);
        }

        //Search button 
        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(FileName))
            {
                ClearAllAData();

                label2.ForeColor = Color.Blue;
                label2.Text = FileName.Split('\\').Last();

                Session = new ProgramSession();
                Session.ParsePage(FileName);
                TbTotalConflictNum.Clear();
                TbTotalConflictNum.Text = Session.Page.NumberofSuspects + "";

                DumpFileWithColors();
                AdjustDisplays(0);
            }
        }

        //Refine 
        private void button8_Click(object sender, EventArgs e)
        {
            Session.Page.RefineSuspects();

            TbTotalConflictNum.Clear();
            TbTotalConflictNum.Text = Session.Page.NumberofSuspects + "";

            AdjustDisplays(0);
        }

        private void ClearAllAData()
        {;
            listBox2.Items.Clear();
            richTextBox1.Clear();
            Session= null;
        }

        private bool IsValidIndex(int index)
        {
            return index >= 0 && index < Session.Page.SuspectLines.Count;
        }

        //VIEW SETTER
        private void AdjustDisplays(int index)
        {
            if (Session.Page.SuspectLines.Count > 0 && IsValidIndex(index))
            {
                TbConflictNumber.Clear();
                TbConflictNumber.Text = (index + 1) + "";

                LbFileView.SetTopIndexAndSelect(Session.Page.SuspectLines[index]);

                label6.Text = @"Conflict View: " + (index + 1);

                LoadRichTxt(Session.Page.Lines[Session.Page.SuspectLines[index]].Line);
            }
        }

        //VIEW SETTER
        private void LoadRichTxt(string line)
        {
            RichDisplayLine.Clear();

            var coloredLine = new LineAnalyzer { Line = line }.BuildColoredLine();
            Session.Wrapper = new Wrapper(coloredLine.PossibleReplacements);

            coloredLine.Write(RichDisplayLine);

            SetWrapDisplay(Session.Wrapper.Any() ? Session.Wrapper.GetCurrent() : "");
        }

        //View Setter
        private void SetWrapDisplay(string word)
        {
            RichDisplayWrap.Clear();
            RichDisplayWrap.Text = word;
            CurrentReplacement.Text = Session.Wrapper.CurrentIndex + 1 + "";
        }

        private void DumpFileWithColors()
        {
            foreach (var aspline in Session.Page.Lines)
            {
                this.listBox2.Items.Add(aspline);
            }
        }

        private void DisplayTxtInBox(TextBox box, string txt)
        {
            box.Clear();
            box.Text = txt;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.AutoSize = true;
            this.label4.Text = @"Conflict View:";
            this.textBox4.Clear();
            this.textBox4.Text = Session.Page.SuspectLines + "";
            this.TotalReplacements.Text = "0";
            this.CurrentReplacement.Text = "0";
        }

        //Next 
        private void button5_Click(object sender, EventArgs e)
        {
            if (Session.SuspectViewIndex + 1 <= Session.Page.SuspectLines.Count)
            {
                Session.SuspectViewIndex++;
                AdjustDisplays(Session.SuspectViewIndex);
            }
        }

        //Previous
        private void button6_Click(object sender, EventArgs e)
        {
            if (Session.SuspectViewIndex > 0)
            {
                Session.SuspectViewIndex--;
                AdjustDisplays(Session.SuspectViewIndex);
            }
        }

        //Wrap
        private void button3_Click(object sender, EventArgs e)
        {
            var wrapValue = Session.WrapAndSave();
            if (wrapValue > 0)
            {
                if (wrapValue == 2)
                {
                    LbChangedView.Items.Add(Session.ChangedLines.Last());
                    this.RichDisplayWrap.Clear();
                    button5_Click(null, null);
                }
                else
                {
                    Session.Wrapper.Next();
                    SetWrapDisplay(Session.Wrapper.GetCurrent());
                }
            }
            else
            {
                button5_Click(null, null);
            }
            
        }


        //Suspect line 
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //CustomReplace
        private void button4_Click(object sender, EventArgs e)
        {

        }

        //Skip 
        private void button7_Click(object sender, EventArgs e)
        {
            if (Session.Wrapper.Next())
            {
                SetWrapDisplay(Session.Wrapper.GetCurrent());
            }
            else
            {
                button5_Click(null, null);
            }
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
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
            if (intval > 0 && intval <= Session.Page.NumberofSuspects)
            {
                Session.SuspectViewIndex = intval - 1;
                AdjustDisplays(Session.SuspectViewIndex);
            }
            RichDisplayLine.Focus();           
        }


        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        //
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

    }

    
}
