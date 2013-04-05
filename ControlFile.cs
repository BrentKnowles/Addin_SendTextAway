// ControlFile.cs
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
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.ComponentModel;
using CoreUtilities;

namespace SendTextAway
{
    /// <summary>
    /// Contains the data (serializable) that influences how the 
    /// automation shoudl behave (templates, et cetera)
    /// </summary>
    public class ControlFile
    {
        public ControlFile()
        {

        }
        const string GENERAL = "General";
        const string EPUB = "Format - ePub";
        const string WORD = "Format - Word";

        private bool _RemoveExcessSpaces=false;
        [Category(GENERAL)]
        [Description("false = leave them, true = remove strings of 5 spaces")]
        public bool RemoveExcessSpaces
        {
            get { return _RemoveExcessSpaces; }
            set { _RemoveExcessSpaces = value; }
        }


        private bool _KeepUnderscore = false;
        [Category(GENERAL)]
        [Description("set to true if you want to keep underscores (no formatting)")]
        public bool UnderscoreKeep
        {
            get { return _KeepUnderscore; }
            set { _KeepUnderscore = value; }
        }

       public enum convertertype { word, epub, text };
        private convertertype _convertype = convertertype.word;
        [Description("What type of conversion to do on the text (word, epub, text)")]
        [Category(GENERAL)]
        public convertertype ConverterType
        {
            get { return _convertype; }
            set { _convertype = value; }
        }

        private int chatmode = 0;
        /// <summary>
        ///  0 = Uunderline
        ///  1 = Italic
        ///  2 = Bold
        /// 
        /// We use this markup (<<<blah<<<) for chat transcrips
        /// </summary>
        [Category(GENERAL)]
        [Description("0 = underline; 1 = italic 2 = bold.")]
        public int ChatMode
        {
            get { return chatmode; }
            set { chatmode = value; }
        }


        private string _description;
        [Category(GENERAL)]
        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        // text message to display at end
        private string _endmessage ="";
        [Category(GENERAL)]
        public string EndMessage
        {
            get { return _endmessage; }
            set { _endmessage = value; }
        }

        private bool mRemoveTabs = false;
        [Description("Smashwords does not want Tabs. This runs a processor removing all tab characters found.")]
        [Category(GENERAL)]
        public bool RemoveTabs
        {
            get { return mRemoveTabs; }
            set { mRemoveTabs = value; }
        }

        private bool mUnderlineShouldBeItalicInstead = false;
        [Description("For Smashwords. Underscore text will be italic instead of underline.")]
        [Category(GENERAL)]
        public bool UnderlineShouldBeItalicInstead
        {
            get { return mUnderlineShouldBeItalicInstead; }
            set { mUnderlineShouldBeItalicInstead = value; }
        }


        private bool mSceneBreakHasTab = true;
        [Description("If set to false there won't be a tab added after scene break (for Smashwords).")]
        [Category(GENERAL)]
        public bool SceneBreakHasTab
        {
            get { return mSceneBreakHasTab; }
            set { mSceneBreakHasTab = value; }
        }
        /// <summary>
        /// i.e., [[~center]]\r\n#\r\n[[~left]]
        /// </summary>
        private string mSceneBreak;
        [Description("Characters to use to indicate a scene break (like #). For HTML based exports this can be a full image path.")]
        [Category(GENERAL)]
        public string SceneBreak
        {
            get { return @mSceneBreak; }
            set { mSceneBreak = @value; }
        }

        private int optionalcode = 1; // bold
        [Description("If you want special formating when the marker [[optional]] is used, i.e., to indicate an optional feature in a game design document indicate 1 for bold or a 2 for italics. A 0 will not format things flaged as optional")]
        [Category(GENERAL)]
        public int OptionalCode
        {
            get { return optionalcode; }
            set { optionalcode = value; }
        }

        private string optional = "(Optional)";
        [Description("Will add this text beside any sentence flagged with the marker [[optional]]")]
        [Category(GENERAL)]
        public string Optional
        {
            get { return optional; }
            set { optional = value; }
        }



        private bool showFootNoteChapter = false;
        [Description("If set to true where the footnote is written will indicate the chapter it appeared in. Usually a debug feature.")]
        [Category(GENERAL)]
        public bool ShowFootNoteChapter
        {
            get { return showFootNoteChapter; }
            set { showFootNoteChapter = value; }
        }
        ///////////////////////////////////////  EPUB


        /// <summary>
        /// ////////////////////////////
        /// </summary>

        #region epubonly
        private string outputdirectory=Constants.BLANK;
        /// <summary>
        /// For things like epub we can specify where the file should go
        /// </summary>
        [Description("Specify where the files should go. Will create a date-named folder therein and a zipped .epub file.")]
        [Category(EPUB)]
        public string OutputDirectory
        {
            get { return outputdirectory; }
            set { outputdirectory = value; }
        }


        private string templatedirectory;
        /// <summary>
        /// For things like epub we can specify where the file should go
        /// </summary>
        [Description("Specify the directory containing source files, footer, and header files, et cetera.")]
        [Category(EPUB)]
        public string TemplateDirectory
        {
            get { return templatedirectory; }
            set { templatedirectory = value; }
        }

        private string zipfile;
        /// <summary>
        /// For things like epub we can specify where the file should go
        /// </summary>
        [Description("Path to .exe that can zip the files after epub creation.")]
        [Category(EPUB)]
        public string Zipper
        {
            get { return zipfile; }
            set { zipfile = value; }
        }
#endregion

        //////////////////////////////////////////////// WORD


        #region WordOnly
        /// <summary>
        /// object reference to the body text style
        /// </summary>
        public object oBodyText; // object reference
        private string bodyText;
        [Description("object reference to the body text style")]
        [Category(WORD)]
        public string BodyText
        {
            get { return bodyText; }
            set
            {
                bodyText = value;
                oBodyText = bodyText;
            }
        }



        
        private bool fancyCharacters = false;
        /// <summary>
        /// Set to true if want ellipsis and quote to be replaced by fancy versions
        /// </summary>
        /// 
        [Description("Set to true if want ellipsis and quote to be replaced by fancy versions")]
        [Category(WORD)]
        public bool FancyCharacters
        {
            get { return fancyCharacters; }
            set { fancyCharacters = value; }
        }

        private string bullet;
        [Description("The style 'List Bullet' is the default")]
        [Category(WORD)]
        public string Bullet
        {
            get { return bullet; }
            set { bullet = value; }
        }

        private string bulletNumber;
        [Description("The style 'List Number' is the default")]
        [Category(WORD)]
        public string BulletNumber
        {
            get { return bulletNumber; }
            set { bulletNumber = value; }
        }







        // UNSORTED BELOW THIS
      

        private string tableStyle;
        /// <summary>
        /// style to use for ||aa|| tables
        /// </summary>
        public string TableStyle
        {
            get { return tableStyle; }
            set { tableStyle = value; }
        }

       /*
        disabled, lost text between page breaks
        private float linespace = 3F;
        [Description("2 = double space")]
        public float Linespace
        {
            get { return linespace; }
            set { linespace = value; }
        }*/
        private string[] fixnumberlist;
        /// <summary>
        /// This is a bit of a hack but you specify second
        /// and third level entrie like
        /// i.
        /// a.
        /// In this type of list
        /// </summary>
        [Description("Help list numbers reset by specifying 2nd, 3rd tier list#s")]
        public string[] FixNumberList
        {
            get { return fixnumberlist; }
            set { fixnumberlist = value; }
        }
        private string heading1;
        public string Heading1
        {
            get { return heading1; }
            set { heading1 = value; }
        }

        private string heading2;
        public string Heading2
        {
            get { return heading2; }
            set { heading2 = value; }
        }


        private string heading3;
        public string Heading3
        {
            get { return heading3; }
            set { heading3 = value; }
        }

        private string heading4;
        public string Heading4
        {
            get { return heading4; }
            set { heading4 = value; }
        }

        private string heading5;
        public string Heading5
        {
            get { return heading5; }
            set { heading5 = value; }
        }


        private string[] multiLineFormats;
        public string[] MultiLineFormats
        {
            get { return multiLineFormats; }
            set { multiLineFormats = value; }
        }
        private string[] multiLineFormatsValues;
        public string[] MultiLineFormatsValues
        {
            get { return multiLineFormatsValues; }
            set { multiLineFormatsValues = value; }
        }


        private string chaptertitle;
        public string ChapterTitle
        {
            get { return chaptertitle; }
            set { chaptertitle = value; }
        }



        private string template;
        /// <summary>
        /// base template to use
        /// </summary>
        public string Template
        {
            get { return template; }
            set { template = value; }
        }

		public static ControlFile Default {
			get {

				ControlFile returnControl = new ControlFile();


				// MAIN CANDIATES
				//*MAJOR*
				returnControl.BodyText = "Body Text Courier";
				returnControl.UnderlineShouldBeItalicInstead = false;
				returnControl.UnderscoreKeep = false;
				returnControl.Template="standardmanuscript.dotx";

				//OTHER STUFF

				returnControl.Bullet = "List Bullet";
				returnControl.BulletNumber = "List Number";
				returnControl.FancyCharacters = false;

				returnControl.ChatMode = 0;
				returnControl.ConverterType = convertertype.word;
				returnControl.Description = "For standard story subs like Analog and Asimov";
				returnControl.EndMessage = "REMEMBER to put space between Address and Title (about 1/3 page)";
				returnControl.Optional = "(Optional)";

				returnControl.OptionalCode = 1;
				returnControl.RemoveExcessSpaces  = false;
				returnControl.RemoveTabs = false;
				returnControl.SceneBreak = "#";
				returnControl.SceneBreakHasTab = true;

				returnControl.ShowFootNoteChapter = false;

				returnControl.ChapterTitle = "Heading Document Top";
				returnControl.FixNumberList = new string[0];

				returnControl.Heading1  = "Heading2 Black Bar";
				returnControl.Heading2 = "Heading3 Lined";
				returnControl.Heading3 = "Heading3 Lined";
				returnControl.Heading4 = "Heading3 Lined";
				returnControl.Heading5 = "Heading4";

				returnControl.MultiLineFormats = new string[4] {"code","quote", "note", "past"};
				returnControl.MultiLineFormatsValues = new string[4] {"Example", "Subtitle", "Subtitle", "bodytext2past_underline"};

				returnControl.TableStyle = "Table Grid 2";

				// build programmatically the 'standard' object



				return returnControl;
			}
		}
        #endregion

    }
}
