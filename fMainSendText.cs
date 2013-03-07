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
                    s = new sendePub();
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

         //TODO: Hook back up
			/* 

            PluginInformationOnTag INFO = ((PluginInformationOnTag)Application.OpenForms[0].Tag);
            string sPlugInParameter = "";
            if (null != INFO)
            {
                INFO.OnGetInfo(1,""); // runs the process
                sPlugInParameter = INFO.OnGetInfo(2,"").ToString();

            }


            //string sPlugInParameter = (string)Application.OpenForms[0].Tag;
            bool parameterfound = true;
            if (sPlugInParameter != null && sPlugInParameter != "")
            {
                string[] sValue = sPlugInParameter.Split(new char[1] { '|' });
                if (sValue.Length != 2)
                {
                    parameterfound = false;
                }
                else
                {
                    // we have two strings (path + preferred control file)
                    //MessageBox.Show(sValue[0]);
                    //MessageBox.Show(sValue[1]);
                    int stopat = 0; // chapter to stop at (dec 11 2010)
                    string sFile = sValue[0];
                    string sControlFile = sValue[1];

                    // but now we pop over to the fPicker and choose another control file
                    fPicker picker = new fPicker(MyPath);
                    picker.ControlFile = sControlFile;
                    if (picker.ShowDialog() == DialogResult.OK)
                    {
                        // rebuild path to file
                        sControlFile = Path.Combine(MyPath, picker.ControlFile);
                        textBoxErrors.Text = Convert(sControlFile, sFile, stopat);// s.WriteText(sFile, sControlFile);


                        if (textBoxErrors.Text == "")
                        {
                            Close();
                        }
                    }
                    else
                    {
                        // if we cancel the Picker we close this main form
                        Close();
                    }



                }
            }
            else
            {
                //we are not a plug-in so initialize the MessageBox
                CoreUtilities.NewMessage.SetupBoxFirstTime(null, "",
      ImageLayout.Stretch,
  Color.AliceBlue, // transprency key  
  new Font("Georgia", 12), new Font("Times", 10),
  Color.Blue,  //button face color
  Color.Black, // caption color
  Color.Black, // message color
  Color.Gray,  // back color for form
  Color.Gray, // back color for caption (turns it into a proper captino heading)
  Color.Gray); // back color for message);

            }



            if (false == parameterfound && Environment.GetCommandLineArgs().Length > 1)
            {
                int stopat = 0; // chapter to stop at (dec 11 2010)

                string sFile = Environment.GetCommandLineArgs()[1];
                string sControlFile = "";
                if (Environment.GetCommandLineArgs().Length == 2)
                {
                   // MessageBox.Show("No control file passed in. Defaulting.");
                    sControlFile = @"e:\controlnormal.xml";
                }
                else
                {
                    sControlFile = Environment.GetCommandLineArgs()[2];
                }


                // allow a length to be passed in?
                if (Environment.GetCommandLineArgs().Length == 4)
                {
                    // stop at a particular chapter number - 1 (i.e., if a 4 passed in we stop at Chapter 3)
                    stopat = Int32.Parse(Environment.GetCommandLineArgs()[3]);
                }
             
               // sendBase s = new sendWord();



                textBoxErrors.Text = Convert(sControlFile, sFile, stopat);// s.WriteText(sFile, sControlFile);


                if (textBoxErrors.Text == "")
                {
                    Close();
                }
            }
            else
            {


                string sDefault = "e:\\controlReilly.xml";
                if (File.Exists(sDefault) == true)
                {
                    Filename = sDefault;
                    propertyGrid1.SelectedObject = (ControlFile)CoreUtilities.General.DeSerialize(
                        Filename, typeof(ControlFile));
                }
            }

*/
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
			/*TODO: hook up
            ADD_Tasks.fTasks tasks = new ADD_Tasks.fTasks();
            tasks.ShowDialog();
            */
        }
    }
}