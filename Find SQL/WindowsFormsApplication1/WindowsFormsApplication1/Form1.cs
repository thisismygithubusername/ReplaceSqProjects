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

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ListBox folders = this.listBox1;
            if (folders.SelectedItem != null)
            {
                RichTextBox outputTextField = this.richTextBox1;
                ListBox matchingLines = this.listBox2;
                List<int> lineNumbers = new List<int>();
                string line = "";
                int counter = 1;

                string fileName = folders.SelectedItem.ToString();
                StreamReader file = new StreamReader(fileName);

                outputTextField.Clear();
                matchingLines.Items.Clear();

                while ((line = file.ReadLine()) != null)
                {
                    if (CheckIfContainsSql(line))
                    {
                        lineNumbers.Add(counter);
                        int length = outputTextField.TextLength;
                        outputTextField.AppendText(line + "\n");
                        outputTextField.SelectionStart = length;
                        outputTextField.SelectionLength = line.Length;
                        outputTextField.SelectionColor = System.Drawing.Color.Red;
                        matchingLines.Items.Add("LINE " + counter + " ---- " + line);
                    }
                    else
                    {
                        outputTextField.AppendText(line + "\n");
                    }
                    counter++;
                }

                file.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog browser = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = browser.ShowDialog();

            string[] files = Directory.GetFiles(browser.SelectedPath);
            ListBox folders = this.listBox1;
            folders.Items.Clear();

            foreach (var file in files)
                folders.Items.Add(file);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ListBox folders = this.listBox1;
            foreach (string file in folders.Items)
            {
                if (file.Contains(this.textBox1.Text))
                {
                    folders.SelectedItem = file;
                    break;
                }
            }
        }

        private bool CheckIfContainsSql(string line)
        {
            if (line.Contains("SELECT") || line.Contains("DELETE") || 
                line.Contains("FROM") || line.Contains("WHERE") || 
                line.Contains("strSQL") || line.Contains("SET"))
                return true;
            return false;
        }
    }
}