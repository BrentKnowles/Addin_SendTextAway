using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using SendTextAway;

namespace ADD_SendTextAway
{
    public partial class fPicker : Form
    {
        private string mFileLocation = "";
        public fPicker(string FileLocation)
        {
            InitializeComponent();
            mFileLocation = FileLocation;

            BuildFileList(FileLocation);
        }
        /// <summary>
        /// populates the listbox
        /// </summary>
        private void BuildFileList(string FileLocation)
        {
            listBoxFiles.Items.Clear();
            foreach (string s in Directory.GetFiles(FileLocation, "*.xml"))
            {
                FileInfo f = new FileInfo(s);
                listBoxFiles.Items.Add(f.Name);
            }
        }


        string mControlFile = "";

        public string ControlFile
        {
            get { return mControlFile; }
            set
            {

                mControlFile = value;
                // feb 15 - removing an issue where there's a crash if no file in SOURCE field
                if (mControlFile != null && mControlFile != "")
                {
                    try
                    {
                        FileInfo f = new FileInfo(mControlFile);

                        for (int i = 0; i < listBoxFiles.Items.Count; i++)
                        {
                            if (listBoxFiles.Items[i].ToString() == f.Name)
                            {
                                // match, go here
                                listBoxFiles.SelectedIndex = i;
                                break;
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        /// <summary>
        /// set the selected item
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            mControlFile = listBoxFiles.Items[listBoxFiles.SelectedIndex].ToString();
            string file = Path.Combine(mFileLocation, mControlFile);
            ControlFile controlFile = (ControlFile)CoreUtilities.FileUtils.DeSerialize(file, typeof(ControlFile));
            if (null != controlFile)
            {
                string imagefile = "";
                FileInfo f = new FileInfo(file);
                string justname = f.Name.Remove(f.Name.IndexOf('.'));
                imagefile = Path.Combine(mFileLocation, justname + ".jpg"); // buidl image file
                pictureBoxPreview.Image = null;
                pictureBoxPreview.Invalidate();
                if (File.Exists(imagefile))
                {
                    pictureBoxPreview.Image = Image.FromFile(imagefile);
                }

                labelType.Text = controlFile.ConverterType.ToString();
                label1.Text = controlFile.Template;
                richTextBoxDetails.Text = controlFile.Description;
            }
        }
    }
}