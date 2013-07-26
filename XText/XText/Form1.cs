using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net.Mail;
using XText;






namespace XPad
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void backgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            richTextBox1.BackColor = colorDialog1.Color; 
        }

        private void fontColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            richTextBox1.ForeColor = colorDialog1.Color; 
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fontDialog1.ShowDialog();
            richTextBox1.Font = fontDialog1.Font;
        }

        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.WordWrap = !(richTextBox1.WordWrap);
            wordWrapToolStripMenuItem.Checked = !wordWrapToolStripMenuItem.Checked;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        protected void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //check for unsaved data
            if (checkUnSavedData() == false)
            {
                e.Cancel = true;
                return;
            }
            SaveSettings(); //Save setting to registry
        }

        private void SaveSettings()
        {
            throw new NotImplementedException();
        }

        private bool checkUnSavedData()
        {
            throw new NotImplementedException();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Open";
            dlg.Filter = "Text Documents(*.txt) |*.txt |All Files(*,*) |*,*";
            dlg.ShowDialog();
            richTextBox1.LoadFile(dlg.FileName);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "Save";
            dlg.Filter = "Text Documents(*.txt) |*.txt |All Files(*,*) |*,*";
            dlg.ShowDialog();
            richTextBox1.SaveFile(dlg.FileName);
        }

        private void pageSetupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.ShowDialog();
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDialog1.ShowDialog();
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Open";
            dlg.Filter = "Text Documents(*.txt) |*.txt |All Files(*,*) |*,*";
            dlg.ShowDialog();
            richTextBox1.LoadFile(dlg.FileName);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "Save";
            dlg.Filter = "Text Documents(*.txt) |*.txt |All Files(*,*) |*,*";
            dlg.ShowDialog();
            richTextBox1.SaveFile(dlg.FileName);
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            richTextBox1.ForeColor = colorDialog1.Color; 
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            fontDialog1.ShowDialog();
            richTextBox1.Font = fontDialog1.Font;
        }

        private void printPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.ShowDialog();
        }

        private void websiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("www.google.com");
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Italic)
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style ^ FontStyle.Italic);

            }

            else
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style | FontStyle.Italic);

            }
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            findToolStripMenuItem.Checked = panel1.Visible = true;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            int index = 0; string temp = richTextBox1.Text; richTextBox1.Text = ""; richTextBox1.Text = temp;

            while (index < richTextBox1.Text.LastIndexOf(textBox1.Text))
            {
                
                richTextBox1.Find(textBox1.Text, index, richTextBox1.TextLength, RichTextBoxFinds.None);
               
                richTextBox1.SelectionBackColor = Color.Orange;
                
                index = richTextBox1.Text.IndexOf(textBox1.Text, index) + 1;
            }
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Bold)
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style ^ FontStyle.Bold);

            }

            else
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style | FontStyle.Bold);

            }
        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Underline)
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style ^ FontStyle.Underline);

            }

            else
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style | FontStyle.Underline);

            }
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 5;
	            richTextBox1.SelectionBullet = true;
	        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void sendTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void viewHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 myform = new Form2();
            myform.ShowDialog();
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            
        }

        private void updatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button3.Visible = true;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            button3.Visible = false;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 myform = new Form3();
            myform.ShowDialog();
        }

        private void leftToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void centerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void rightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void noneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 0;
        }

        private void ptsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 5;
        }

        private void ptsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 10;
        }

        private void ptsToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 15;
        }

        private void ptsToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 20;
        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Underline)
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style ^ FontStyle.Bold);

            }

            else
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style | FontStyle.Bold);

            }
        }

        private void italicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Underline)
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style ^ FontStyle.Italic);

            }

            else
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style | FontStyle.Italic);

            }
        }

        private void undelineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Underline)
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style ^ FontStyle.Underline);

            }

            else
            {

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, richTextBox1.SelectionFont.Style | FontStyle.Underline);

            }
        }

        private void bulletsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionIndent = 5;
            richTextBox1.SelectionBullet = true;
        }
        }
       


    }
