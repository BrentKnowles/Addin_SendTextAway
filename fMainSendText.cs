// fMainSendText.cs
//
// Copyright (c) 2013 Brent Knowles (http://www.brentknowles.com)
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
// Review documentation at http://www.yourothermind.com for updated implementation notes, license updates
// or other general information/
// 
// Author information available at http://www.brentknowles.com or http://www.amazon.com/Brent-Knowles/e/B0035WW7OW
// Full source code: https://github.com/BrentKnowles/YourOtherMind
//###
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

using Word = NetOffice.WordApi;
//using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


/*
 * Intended as a command-line program invoked to take a body of text written in "WIKI-format" 
 * and convert it to Word Doc
 * 
 * Stages
 * 1. Basic text writing
 * 2. Selecting a style sheeting in the opening
 * 3. Converting wiki tags (= header 1 =, et cetera) into translated foramtting
 * 4. Allow the user to override what formatting does? (or should that just be purely based on the style selected?)
 * 
 */
namespace SendTextAway
{
    public partial class fMain : Form
    {
        
        public fMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
        }
        /// <summary>
        /// Other Control MEthods TO Implement
        /// - control over paragraph spacing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (filename == "")
            {
                MessageBox.Show("You must open a control file first");
                return;
            }
            SaveCurrentControlFile();

            string sFile = Path.GetTempFileName();
            StreamWriter writer = new StreamWriter(sFile);

            for (int i = 0; i < richTextBox1.Text.Length; i++)
            {
                writer.Write(richTextBox1.Text[i]);
            }
            
            writer.Flush();
            writer.Close();

            Convert(filename, sFile, 0);
            


           
        }

        /// <summary>
        /// wrapper for doing an actual converter
        /// </summary>
        public string Convert(string sControlFile, string sFile, int stopat)
        {
            ControlFile zcontrolFile = (ControlFile)CoreUtilities.FileUtils.DeSerialize(sControlFile, typeof(ControlFile));
            // may 2010 adding ablity to do epub files too
            string sError = "";
            if (null != zcontrolFile)
            {
                sendBase s = null;
                if (zcontrolFile.ConverterType == ControlFile.convertertype.word)
                {
                    s = new sendWord();
                }
                else if (zcontrolFile.ConverterType == ControlFile.convertertype.epub)
                {
                    s = new sendePub2();
                }
                 sError = s.WriteText(sFile, zcontrolFile, stopat);
                textBoxErrors.Text = sError;
            }
            return sError;
        }

        private string filename = "";
        private string Filename
        {
            get { return filename; }
            set { filename = value;
            this.Text = filename;
            }
        }

        private void saveAsControlFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = MyPath;
            save.DefaultExt = "xml";
            save.Filter = "Control Files|*.xml";
            if (save.ShowDialog() == DialogResult.OK)
            {
              /*  ControlFile controlFile = new ControlFile();
                controlFile.Template = "ORA.dot";
                controlFile.BodyText = "Body Text,b";
                controlFile.ChapterTitle = "ChapterTitle,ct";
                controlFile.Heading1 = "Heading 1";
                controlFile.Heading2 = "Heading 2";*/
                Filename = save.FileName;
                SaveCurrentControlFile();
            }
        }

        /// <summary>
        /// path to look for control files
        /// </summary>
        public string MyPath
        {
            get
            {

                string path = @"C:\Users\BrentK\Documents\Keeper\SendTextAwayControlFiles";
                if (Directory.Exists(path) != true)
                {
                    path = Application.StartupPath;
                }
                return path;
            }
        }

        private void openControlFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory = MyPath;
            open.DefaultExt = "xml";
            open.Filter = "Control Files|*.xml";
            if (open.ShowDialog() == DialogResult.OK)
            {
                Filename = open.FileName;
            propertyGrid1.SelectedObject = (ControlFile)CoreUtilities.FileUtils.DeSerialize(
                open.FileName, typeof(ControlFile));
            }
        }
        private void SaveCurrentControlFile()
        {
            if (Filename != "")
            {
                CoreUtilities.FileUtils.Serialize(propertyGrid1.SelectedObject, Filename,"");
            }
        }
        private void saveControlFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveCurrentControlFile();
            
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            propertyGrid1.SelectedObject = new ControlFile();
        }

        private void fMain_Load(object sender, EventArgs e)
        {

       
        }

        private void onlyToChapter2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename == "")
            {
                MessageBox.Show("You must open a control file first");
                return;
            }
            SaveCurrentControlFile();

            string sFile = Path.GetTempFileName();
            StreamWriter writer = new StreamWriter(sFile);

            for (int i = 0; i < richTextBox1.Text.Length; i++)
            {
                writer.Write(richTextBox1.Text[i]);
            }

            writer.Flush();
            writer.Close();

            Convert(filename, sFile, 2);
            
        }

        private void tasksToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}