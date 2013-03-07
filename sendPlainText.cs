using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;

namespace SendTextAway
{
    /// <summary>
    /// probably never used but exports a basic text file
    /// </summary>
    public class sendPlainText : sendBase
    {
        protected StreamWriter file1;
        protected string currentFileBeingWritten = "";

        protected ArrayList files; // this will be used to list the files for the TOC (by default will also include legal and copyright

        protected override int InitializeDocument(ControlFile _controlFile)
        {
            files = new ArrayList();
            base.InitializeDocument(_controlFile);

            // if we have not already setup a filename, we create one now
            if ("" == currentFileBeingWritten)
            {
                currentFileBeingWritten = Path.Combine(_controlFile.OutputDirectory, "textoutput_" + DateTime.Today.ToString("MM yy") + ".txt");
            }

            StartNewFile(currentFileBeingWritten);

			return 1;
           

        }

        

        /// <summary>
        /// inserts text for header file
        /// </summary>
        protected virtual void InsertFooterorNewFile(string sChapterToken)
        {
        }
        /// <summary>
        /// inserts text for header file
        /// </summary>
        protected virtual void InsertHeaderForNewFile(string sChapterToken)
        {
        }

        /// <summary>
        ///  anew file has been started
        /// </summary>
        /// <param name="sFile"></param>
        protected void StartNewFile(string sFile)
        {
            if (File.Exists(sFile) == true)
            {
                File.Delete(sFile);
            }
            file1 = new StreamWriter(sFile);

            if (sFile.IndexOf("preface") > -1)
            {
                InsertHeaderForNewFile("Preface");
            }
            else
            if (sFile.IndexOf("footnote") > -1)
            {
                InsertHeaderForNewFile("Footnotes");
            }
            else
                InsertHeaderForNewFile(chaptertoken);
            FileInfo file = new FileInfo(sFile);
            files.Add(file.Name);
            file = null;
            
            
        }

        /// <summary>
        /// While processing linline formatting (bold, et cetera), this is used to write out a line of text
        /// </summary>
        /// <param name="sText"></param>
        protected override void InlineWrite(string sText)
        {
            file1.WriteLine(sText);
            //oSelection.TypeText(sText);
        }
        /// <summary>
        /// with children we will be opening and closing multiple fiels (i.e., for each chapter)
        /// This can be called directlyt o close them
        /// </summary>
        protected virtual void CloseCurrentFile()
        {
            InsertFooterorNewFile(chaptertoken);
            file1.Close();
            file1.Dispose();
        }


        /// <summary>
        /// at end?
        /// </summary>
        protected override void Cleanup ()
		{
			base.Cleanup ();
			CloseCurrentFile ();
			if (false == SuppressMessages) {
				CoreUtilities.NewMessage.Show (CoreUtilities.Loc.Instance.GetStringFmt ("{0} Has been written out", currentFileBeingWritten));
			}
        }
		public override string ToString ()
		{
			return string.Format ("[sendPlainText]");
		}
    }
}
