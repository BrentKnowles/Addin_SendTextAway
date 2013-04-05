// sendBase.cs
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
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;

using CoreUtilities;
using Word = NetOffice.WordApi;
// remove this
//using Word = Microsoft.Office.Interop.Word;
namespace SendTextAway
{
   
    /// <summary>
    /// base object
    /// 
    /// Word and Blog writers will descend from this.
    /// </summary>
    public class sendBase
    {

        protected string FootnoteLog = ""; // this is just a list of all footnotes included 
        protected System.Collections.Hashtable tallyTable;
        protected System.Collections.Hashtable variableTable;
        protected bool InAList = false;
        protected int chapter = 1;
        protected Hashtable FootnoteHash;
        protected Hashtable Footnotes_in_specific_chapter;
		// added for unit testing
		public bool SuppressMessages = false;

        protected string chaptertoken
        {
            get { return "Chapter " + chapter.ToString(); }

        }


         protected ControlFile controlFile = null;

        protected string lastLink = ""; // when we encounter |www.|me link| we store the www.
                                        // and then when the next | is hit we grab it and build
                                        // and proper link with the label 'me link' 
        private string sError = ""; //holder error information
        
        protected string sFulltext = "";

        protected void UpdateErrors(string sNewInfo)
        {
            sError = sError + "\n" + sNewInfo;
        }

        protected System.Collections.ArrayList bookmarks = null; // contain a list of links to cross reference in cleanup
        protected int _stopat = 0;

        /// <summary>
        /// This routine will handle the parsing of the file and will call other functions
        /// that will be INHERITED by the child classes to actually handle the creation of the end documentation
        /// </summary>
        /// <param name="sTextFile">The plain text file, in wiki-code format, to parse</param>
        /// <param name="sControlFile">a file containing info on what template should be used and such</param>
        /// <param name="stopat">the chapter -1 to stop at (4 means stop at end of chapter 3)</param>
        public string WriteText(string sTextFile, /*string sControlFile*/ ControlFile zcontrolfile, int stopat)
        {
            bookmarks = new System.Collections.ArrayList();
            _stopat = stopat; // chapter to stop at
            int linecount = 0;
            int newlinecount = 0;
            int blankline = 0;
            if (File.Exists(sTextFile) == false)
            {
                throw new Exception(sTextFile + " does not exist");
            }
            if (null == zcontrolfile)
            {
                throw new Exception("Control file does not exist");
            }
           // try
          //  {
              //  ControlFile zcontrolFile = (ControlFile)CoreUtilities.General.DeSerialize(sControlFile, typeof(ControlFile));
               if ( InitializeDocument(zcontrolfile) == -1) return Constants.BLANK;


                // load template file (ToDo: must do this for real eventually)


                StreamReader reader = new StreamReader(sTextFile);
               
                
                string sLine = reader.ReadLine();
                while (sLine != null)
                {
                    linecount++; // may 2012 tracking line count
                    if ("" == sLine)
                    {
                        blankline++;
                    }
                    string sText = sLine;
                    sFulltext = sFulltext + sLine; // fulltext is simply used by some routines for error checking

                   /* if (sText.IndexOf("Motto") > -1)
                    {
                        int ii = 9;
                    }
*/

                    // if we have hit chapter threshold than stop
                    if ((stopat > 0) && chapter == (stopat + 1))
                    {
                        break;
                    }


                    // January 2012
                    // We have an issue when saving as HTML that my fancy heading text
                    // becomes really ugly UNDERLINES
                    // I'm thinking of stripping out long rows of spaces
                    if (controlFile.RemoveExcessSpaces == true)
                    {
                        while (sText.IndexOf("     ") > -1)
                        {
                            sText = sText.Replace("     ", "");
                        }
                    }


                    if (sText.IndexOf("---") > -1)
                    {
                        NewMessage.Show("You have three hyphens. Probably error. Quitting");
                        return "";
                    }

                    if (sText.IndexOf("\"'") > -1)
                    {
                        throw new Exception("You have a \" mixed with a ' which is likely an error. Shutting down.");
                     }
                     if (sText.IndexOf("'\"") > -1)
                     {
                         throw new Exception("You have a \" mixed with a ' which is likely an error. Shutting down.");
                     }

                    // log common typos
                    if (sText.IndexOf("\"'") > -1)
                    {
                        UpdateErrors("Possible Errors: " + sText);
                    }

                    // first test for table
                    if (sText.IndexOf("||") == 0)
                    {
                        AddTable(sText);
                    }
                    else
                    FormatRestOfText(sText);


              

                    sLine = reader.ReadLine();
                }
                //oSelection.ParagraphFormat.LineSpacing = 2F;
                // oSelection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle;
                reader.Close();

                newlinecount = this.LineCount();
                Cleanup();
         /*   }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }*/

            string line = String.Format("LINE COUNT OLD: {0} BLANK: {1} NEW: {2}",linecount.ToString(),
                blankline, newlinecount.ToString());
            sError = sError + Environment.NewLine + line;

            if (newlinecount != (linecount+blankline))
            {
                string sLineError = String.Format("LINES DO NOT MATCH. Original lines {0} + blank lines {1} = {2} but new lines is = {3}. NOTE: Extra blank space at TOP of will throw count off.",
                    linecount, blankline, linecount + blankline, newlinecount);

				if (false == SuppressMessages)
				{
                NewMessage.Show(sLineError);
				}
            }
            if (error_CountInlineBrackets > 0)
            {
                string ExtraBrackets = String.Format("You have extra brackets #{0}", error_CountInlineBrackets);
                NewMessage.Show(ExtraBrackets);
            }

            // also add it to END message
            controlFile.EndMessage = String.Format("{0} {1} {2}", controlFile.EndMessage,
                Environment.NewLine, line);
            return sError;
        }
        protected virtual void SearchFor(object searchMe)
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected virtual int LineCount()
        {
            return 0;
        }

        protected string sLastTitle = "";

        protected void FormatRestOfText(string sText)
        {
            FormatRestOfText(sText, true);
        }
        /// <summary>
        /// formats the rest of the text, this was pulled out so that Tables can call formatting
        /// </summary>
        /// <param name="sText"></param>
        protected void FormatRestOfText(string sText, bool AddLineFeed)
        {
            // if there are tab characters in the text we remove them
            if (true == controlFile.RemoveTabs)
            {
                sText = sText.Replace("\t", "");
            }

            try
            {

                // dec 2009 - only considerit a bullet if there is somet text here


                
                int hashIndex = -1;
                try
                {
                    hashIndex = sText.IndexOf("#");
                }
                catch (Exception ex)
                {
                    NewMessage.Show(ex.ToString());
                }

                int asterixIndex = -1;
                try
                {

                    asterixIndex = sText.IndexOf("*");
                }
                catch (Exception ex)
                {
                    NewMessage.Show(ex.ToString());
                }


                if (hashIndex == 0 && sText.Length > 1)
                {
                    AddNumberedBullets(sText);

                }
                else
                    if (asterixIndex == 0)
                    {
                        AddBullets(sText);

                    }
                    else
                        if (IsFormat("<t>", "</t>", ref sText, false) == true)
                        {
                            AddTitle(sText);
                            sLastTitle = sText; // we store this to use in [[f]] type indexes



                        }
                        else
                            if (IsFormat("=====", "=====", ref sText, true) == true)
                            {
                                AddHeader(sText, 5);



                            }
                            else
                                if (IsFormat("====", "====", ref sText, true) == true)
                                {
                                    AddHeader(sText, 4);



                                }
                                else
                                    if (IsFormat("===", "===", ref sText, true) == true)
                                    {

                                        AddHeader(sText, 3);


                                    }

                                    else
                                        if (IsFormat("==", "==", ref sText, true) == true)
                                        {

                                            AddHeader(sText, 2);


                                        }
                                        else
                                            if (IsFormat("=", "=", ref sText, true) == true)
                                            {

                                                AddHeader(sText, 1);


                                            }
                                            else
                                            {

                                                // go through each multile line building up end
                                                // and begin tags
                                                bool bMultiChange = false;
                                                for (int counter = 0; counter < controlFile.MultiLineFormats.Length; counter++)
                                                {

                                                    if (sText == String.Format("<{0}>", controlFile.MultiLineFormats[counter]))
                                                    {
                                                        bMultiChange = AddStartMultiLineFormat(controlFile.MultiLineFormatsValues[counter]);

                                                    }
                                                    else
                                                        if (sText == String.Format("</{0}>", controlFile.MultiLineFormats[counter]))
                                                        {
                                                            bMultiChange = true;
                                                            AddEndMultiLineFormat();

                                                        }
                                                        else // dec 11 2010 - sometimes /t controls sneak in, strip em
                                                        {
                                                            string sTemp = sText.Replace("/t", " ").Trim();
                                                            if ( sTemp == String.Format("</{0}>", controlFile.MultiLineFormats[counter]))
                                                            {
                                                                bMultiChange = true;
                                                                AddEndMultiLineFormat();
                                                            }
                                                           
                                                        }
                                                }


                                                if (bMultiChange == false)
                                                {

                                                    ApplyAllInlineFormats(sText, controlFile, AddLineFeed);
                                                }
                                                // oSelection.TypeText(sText + Environment.NewLine);
                                            }
            }
            catch (Exception ex)
            {
                CoreUtilities.NewMessage.Show(ex.ToString());
            }
        }
     //   string TEXT_NULL = "++NULL++";
        /// <summary>
        /// used for replacing fancy quotes and the like
        /// </summary>
        /// <param name="sSource"></param>
        /// <returns></returns>
        protected string ReplaceFancyCharacters(string sSource)
        {
            if (true == controlFile.FancyCharacters)
            {
                
                sSource = sSource.Replace(".\"", ".\x201D");

                sSource = sSource.Replace("\".", "\x201D.");
            

                sSource = sSource.Replace("!\"", "!\x201D");
                sSource = sSource.Replace("?\"", "?\x201D");
                sSource = sSource.Replace("--\"", "--\x201D");

                sSource = sSource.Replace("-\"", "-\x201D");

                sSource = sSource.Replace("-- \"", "-- \x201D");
                sSource = sSource.Replace(",\"", ",\x201D");
                // remainder of quotes will be the opposite way
                sSource = sSource.Replace("\"", "\x201C");

                // do standard repalcements (February 2010)
                if (sSource.IndexOf("...") > -1)
                {
                    sSource = sSource.Replace("...", "\x2026");
                }

                // to finish look here: http://www.unicode.org/charts/charindex.html
            }
            return sSource;
        }

        protected void ApplyAllInlineFormats(string sSource, ControlFile controlFile)
        {
            ApplyAllInlineFormats(sSource, controlFile, true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sSource"></param>
        /// <param name="oSelection"></param>
        /// <param name="controlFile"></param>
        protected void ApplyAllInlineFormats(string sSource, ControlFile controlFile, bool AddLineFeed)
        {
            string sOriginal = sSource;
            string[] sFormats = new string[10];
            sFormats[0] = "'''''";
            sFormats[1] = "'''";
            sFormats[2] = "''";
            sFormats[3] = "_";
            sFormats[4] = "---";
            sFormats[5] = "|"; // links
            sFormats[6] = "[[";
            sFormats[7] = "^"; // O^2^
            sFormats[8] = "##";
            sFormats[9] = "<<<";
          //  sFormats[9] = "..."; can't be used for non doubled , won't work

            // Ok, regex not working with mulitple matches on a line
            // we do this the old fashioned way

            // we look for starttag matches
            // and then for endtag matches

            // we cycle through each format

            string sFrontTag = "";
            string sEndTag = "";
            int nFormat = 0;


            // we cycle through the formats. Once we find one format, we apply it
            // trimming the substring until all is done
            int nFirstStart = 0;

          

            sSource = ReplaceFancyCharacters(sSource);


            int nLastStart = -1;
            int nLastPos = 0;

            while (nFirstStart > -1)
            {

                nFormat = 0;
                nFirstStart = -1;
                // find a format
                while (nFormat < sFormats.Length)
                {
                    // we are looking for the closest match to start
                    // our processing at the beginning of the string
                    int nPossible = sSource.IndexOf(sFormats[nFormat]);
                    if ((nPossible < nFirstStart || nFirstStart == -1) && nPossible != -1)
                    {

                        sFrontTag = sFormats[nFormat];
                        sEndTag = sFrontTag;
                        if (nFormat == 6)
                        {
                            // special case, end tag must be end ]]
                            sEndTag = "]]";
                        }
                        nFirstStart = nPossible;
                    }
                    nFormat++;

                    // nFirstStart = nPossible;
                }
                if (nFirstStart == -1)
                {
                    // we found no format matches
                    // this means it is time to exit
                    break;
                }

                // write text from lastpos to now
                string sFirst = sSource.Substring(nLastPos, nFirstStart);



                try
                {
                    nLastStart = sSource.IndexOf(sEndTag, nFirstStart + sFrontTag.Length);
                }
                catch (Exception ex)
                {
                    NewMessage.Show(ex.ToString());
                }

                if (nLastStart > -1)
                {
                    InlineWrite(sFirst);
                }

                if (nLastStart > -1)
                {

                     string sBoldText = "";

                    // We only complete if there is an end tag (this is to give rid of an info loop --
                    if (nLastStart > -1)
                    {
                        sBoldText = sSource.Substring(nFirstStart + sFrontTag.Length,
                             nLastStart - nFirstStart - sFrontTag.Length);

                        //oSelection.TypeText(sBoldText);
                       


                        sSource = sSource.Substring(nLastStart + sEndTag.Length, sSource.Length - nLastStart - sEndTag.Length);
                    }

                    if (sFrontTag == "'''''")
                    {
                        InlineBold(1);
                        InlineItalic(1);


                    }
                    else
                        if (sFrontTag == "'''")
                        {

                            InlineBold(1);
                            //   object oStrong = "Strong, fb";
                            //   oSelection.set_Style(ref oStrong);
                        }
                        else
                            if (sFrontTag == "''")
                            {
                                InlineItalic(1);
                            }
                            else
                                if (sFrontTag == "_")
                                {
                                    
                                    if (false == controlFile.UnderscoreKeep)
                                    {
                                        if (true == controlFile.UnderlineShouldBeItalicInstead)
                                        {
                                            InlineItalic(1);
                                        }
                                        else
                                            InlineUnderline(1);
                                    }
                                    else
                                    {
                                        // don't underline, keep underscores
                                        InlineWrite("_");
                                    }
                                }
                                
                                else if (sFrontTag == "^")
                                {
                                    InlineSuper(1);
                                }
                                else if (sFrontTag == "##")
                                {
                                    InlineSub(1);
                                }
                                    else

                                    if (sFrontTag == "<<<")
                                    {
                                        ChatMode(1);
                                    }

                                    else

                                        if (sFrontTag == "---")
                                        {
                                            if (nLastStart > -1)
                                            {
                                                InlineStrikeThrough(1);
                                            }
                                        }
                                        else

                                            if (sFrontTag == "|")
                                            {
                                                /* object oTextToShow = ("me a link");
                                                 object oLink = "http://www.google.com";
                                                 oSelection.Hyperlinks.Add(oSelection.Range, ref oLink, ref  oMissing,
                                                    ref  oMissing, ref oTextToShow, ref oMissing);*/
                                                if (lastLink != "")
                                                {
                                                    AddLink(lastLink, sBoldText);
                                                    lastLink = "";
                                                    sBoldText = "";
                                                }
                                                else
                                                    if (sBoldText.IndexOf("##") > -1 ||
                                                        sBoldText.IndexOf("www.") > -1 ||
                                                        sBoldText.IndexOf("http:") > -1)
                                                    {
                                                        // error checking
                                                        // look for broken bookmarks

                                                        if (sBoldText.IndexOf("##") > -1)
                                                        {
                                                            bookmarks.Add(sBoldText.Replace("##", "").Trim());
                                                        }

                                                        lastLink = sBoldText;


                                                        sBoldText = "";
                                                    }


                                                    else
                                                        if (CoreUtilities.General.IsGraphicFile(sBoldText) == true)
                                                        {
                                                            AddPicture(sBoldText);
                                                            sBoldText = ""; // blank thte text
                                                        }
                                            }
                                            else if (sFrontTag == "[[")
                                            {
                                                bool bOptional = false;
                                                bool bAllowWrite = false;

                                                if (sBoldText.IndexOf("~center") > -1)
                                                {
                                                    AlignText(0);
                                                    sBoldText = ""; /// June 2011 process remainderFAILED, GAVE UP (basically I could not have multiple codes on one line)
                                                }
                                                else
                                                    if (sBoldText.IndexOf("~left") > -1)
                                                    {
                                                        AlignText(1);
                                                        sBoldText = ""; /// June 2011 process remainderFAILED, GAVE UP
                                                     //   sBoldText = Environment.NewLine;
                                                    }
                                                    else
                                                        if (sBoldText.IndexOf("~right") > -1)
                                                        {
                                                            AlignText(2);
                                                            sBoldText = ""; /// June 2011 process remainder FAILED, GAVE UP
                                                        }
                                                        else // failed gave up
                                                // June 2 2011 - intentionally removed the else here
                                                // because these formats should work with others
                                                        
                                                            if (sBoldText.IndexOf("~pagebreak") > -1)
                                                            {
                                                                AddPageBreak();
                                                                sBoldText = "";
                                                            }
                                                            else
                                                                if (sBoldText.IndexOf("~break") > -1)
                                                                {
                                                                    InAList = false;
                                                                    sBoldText = "";
                                                                }
                                                                else /// failed gave up

                                                            // June 2 2011 - intentionally removed the else here
                                                            // because these formats should work with others

                                                                if (sBoldText.IndexOf("~id") > -1)
                                                                {
                                                                    // id is used by sendePub to parse things like Title information
                                                                    // so that gets output into the generated epub book
                                                                    sBoldText = sBoldText.Replace("~id", "");
                                                                    string[] idparts = sBoldText.Split(new char[1] { '|' });
                                                                    if (idparts != null && idparts.Length == 2)
                                                                    {
                                                                        AddId(idparts[0].Trim(), idparts[1].Trim());
                                                                    }

                                                                    sBoldText = "";
                                                                }
                                                                else
                                                                    // we found a footnote and are adding it to the HashTable (no outputting)
                                                                    if (sBoldText.IndexOf("footnoteadd") > -1)
                                                                    {
                                                                        sBoldText = sBoldText.Replace("footnoteadd", "");
                                                                        string[] footparts = sBoldText.Split(new char[1] { '|' });
                                                                        AddFootnoteText(footparts[0].Trim(), footparts[1].Trim());
                                                                        sBoldText = "";
                                                                    }
                                                                    else //output the LINK (and in some case actual footnote)
                                                                        if (sBoldText.IndexOf("footnotelink") > -1)
                                                                        {
                                                                            /* Implementation
                                                                             * All footnotes in the 'first file' processed, which is not written. 
                                                                             * Instead the notes are stored in a HASH (FootnoteHash).
                                                                             * Now as we hit FOOTNOTELINK we lookup the footnote from the hash
                                                                             */

                                                                            InlineWrite(" "); // August 2010 insert a blank space to make footnote work?

                                                                            sBoldText = sBoldText.Replace("footnotelink", "");
                                                                            // the text remaining should be the name of the footnotelink

                                                                            AddFootnote(sBoldText.ToLower());

                                                                            // store the number of footnotes in the current chapter
                                                                            if (Footnotes_in_specific_chapter.ContainsKey(chapter) == true)
                                                                            {
                                                                                Footnotes_in_specific_chapter[chapter] = ((int)Footnotes_in_specific_chapter[chapter]) + 1;


                                                                            }
                                                                            else
                                                                            {
                                                                                Footnotes_in_specific_chapter.Add(chapter, 1);
                                                                            }


                                                                            FootnoteLog = FootnoteLog + "," + sBoldText.ToLower();
                                                                            sBoldText = "";

                                                                        }
                                                                        else
                                                                            if (sBoldText.IndexOf("Anchor") > -1)
                                                                            {
                                                                                sBoldText = sBoldText.Replace("Anchor", "");
                                                                                AddBookmark(sBoldText);
                                                                                sBoldText = "";
                                                                            }
                                                                           // else if (sBoldText.IndexOf("~scene") > -1)
                                                                            else if (YomParse.KeywordMatch(sBoldText, YomParse.KEYWORD_SCENE) == true)
                                                                            {
                                                                                // if we find a break we pull the macro that has been assigned 
                                                                                // to the properties ControlFile and put that text here
                                                                                // F 2010 - I use this for typing less text with my scene breaks
                                                                                //.InlineWrite("[[~center]]#[[~left]]");
                                                                                //string sResult = @controlFile.SceneBreak; //[[~center]]\r\n#\r\n[[~left]]
                                                                                //@sResult = @sResult.Replace(@"\\", @"\");
                                                                                //Console.WriteLine(@sResult);
                                                                                //InlineWrite("\r\n");
                                                                                FormatRestOfText("[[~center]]", false);
                                                                                FormatRestOfText(controlFile.SceneBreak, true);
                                                                                FormatRestOfText("[[~left]]", false);
                                                                                if (true == controlFile.SceneBreakHasTab)
                                                                                {
                                                                                    AddTab();
                                                                                    // February 15 2011
                                                                                    // The other fix I made to keep
                                                                                    // from losing NEEDED LINESPACE
                                                                                    // unfortuantely is now adding
                                                                                    // linespace after a scene.
                                                                                    //Trying to remove this.
                                                                                    AddLineFeed = false;
                                                                                }
                                                                                //sBoldText = Environment.NewLine + Environment.NewLine + sBoldText;
                                                                                // FormatRestOfText(@"[[~center]]\r\n#\r\n[[~left]]");
                                                                            }
                                                                            else
                                                                                if (sBoldText.IndexOf("~TOC") > -1)
                                                                                {
                                                                                    AddTableOfContents();
                                                                                    sBoldText = "";
                                                                                }
                                                                                else
                                                                                    if (sBoldText.IndexOf("~var=") > -1)
                                                                                    {
                                                                                        VariableAdd(sBoldText);
                                                                                        sBoldText = "";
                                                                                    }
                                                                                    else
                                                                                        if (sBoldText.IndexOf("~var") > -1)
                                                                                        {
                                                                                            // swap in variable

                                                                                            sBoldText = VariableSwap(sBoldText);

                                                                                            // need to solve linefeed issue

                                                                                            ApplyAllInlineFormats(sBoldText, controlFile, false);
                                                                                            sBoldText = " "; // inteintiona July 2011
                                                                                            if (sBoldText != "")
                                                                                            {
                                                                                                bAllowWrite = true;
                                                                                            }
                                                                                        }
                                                                                        else
                                                                                            if (sBoldText.IndexOf("~Footnotecount") > -1)
                                                                                            {
                                                                                                sBoldText = sBoldText.Replace("~Footnotecount", "").Trim();

                                                                                                if (Footnotes_in_specific_chapter.Count > 0)
                                                                                                {
                                                                                                    ArrayList list = new ArrayList(Footnotes_in_specific_chapter.Keys);



                                                                                                    list.Sort();
                                                                                                    foreach (int key in list)
                                                                                                    {
                                                                                                        AddTable((key - 1) + " " + Footnotes_in_specific_chapter[key]);

                                                                                                    }


                                                                                                }

                                                                                            }
                                                                                            else
                                                                                                if (sBoldText.IndexOf("~Display") > -1)
                                                                                                {
                                                                                                    // System.Windows.Forms.MessageBox.Show(tallyTable.Count.ToString());
                                                                                                    sBoldText = sBoldText.Replace("~Display", "").Trim();
                                                                                                    // get just source like b
                                                                                                    if (tallyTable.ContainsKey(sBoldText) == true)
                                                                                                    {
                                                                                                        //InlineWrite(tallyTable[sBoldText].ToString());
                                                                                                        string[] displayitems = tallyTable[sBoldText].ToString().Split(',');
                                                                                                        if (displayitems != null)
                                                                                                        {
                                                                                                            AddTable("||Reference||\n");
                                                                                                            foreach (string s in displayitems)
                                                                                                            {

                                                                                                                AddTable(String.Format("{0}\n", s));
                                                                                                                //ApplyAllInlineFormats(s, controlFile);
                                                                                                                //InlineWrite( s  + "\n");
                                                                                                            }
                                                                                                        }
                                                                                                    }

                                                                                                }
                                                                                                else
                                                                                                    if (sBoldText.IndexOf("~Tally") > -1)
                                                                                                    {
                                                                                                        sBoldText = sBoldText.Replace("~Tally", "").Trim();
                                                                                                        // get just source like b
                                                                                                        if (tallyTable.ContainsKey(sBoldText) == true)
                                                                                                        {
                                                                                                            InlineWrite(tallyTable[sBoldText].ToString().Split(',').Length.ToString());
                                                                                                        }
                                                                                                    }
                                                                                                    else
                                                                                                    {

                                                                                                        if (sBoldText.IndexOf("optional") > -1)
                                                                                                        {
                                                                                                            bOptional = true;
                                                                                                        }
                                                                                                        //  VariableCheck(sBoldText, sOriginal, sFormats);

                                                                                                        // in this situation we intentional DO NOT take
                                                                                                        // the return value from VariableCheck unlike
                                                                                                        /// the later call which is intended for when 
                                                                                                        /// we have code within formatting
                                                                                                        
                                                                                                        // may 2012 
                                                                                                        VariableCheck(sBoldText, sOriginal, sFormats, false);
                                                                                                      //  VariableCheck2(sOriginal, sFormats);
                                                                                                    }
                                                if (bOptional == true)
                                                {
                                                    // optinal is special, we insert formatting expression
                                                    switch (controlFile.OptionalCode)
                                                    {
                                                        case 1: InlineBold(1); break;
                                                        case 2: InlineItalic(1); break;
                                                    }

                                                    sBoldText = " " + controlFile.Optional;
                                                }
                                                else
                                                    if (bAllowWrite == false)
                                                    {
                                                        sBoldText = ""; // we don't want to write codes
                                                      //  AddLineFeed = false; // june 2011 to get rid of empyt comments

                                                        //January 2012
                                                        // I am missing some linefeeds, specifically when a [[f]] is
                                                        // anywhere in preceding line.
                                                        // I think I need to test to see if a linefeed is PRESENT
                                                        // in sOriginal
                                                        // DID NOT WORK! So I commented out the June 2011 change
                                                        /*
                                                        if (sOriginal.IndexOf("\\n") > -1 || sOriginal.IndexOf("\\r") > -1)
                                                        {
                                                            AddLineFeed = true;
                                                        }*/
                                                    }
                                            }

                    // june 2010 moving things here so formatted text can be parsed for this
                    if (sBoldText.IndexOf("[[") > -1)
                    {
                        // this applies ONLY when we have code with formatting.
                        
                       sBoldText =  VariableCheck(sBoldText, sOriginal, sFormats, true);

                        

                    }
                  
                    //if (TEXT_NULL != sBoldText) // June 2 2011 -- don't write blank lines from Varisble comments
                    {
                        InlineWrite(sBoldText);
                    }
                    /*OVerkill   else
                       {
                           // we had only a start tag and no end, so we write out the offending
                           // text (the FrontTag) and then chop it off.
                           InlineWrite(sFrontTag);
                           char[] trimChars = new char[sFrontTag.Length];
                           for (int i= 0; i < sFrontTag.Length; i++)
                           {
                               char c = sFrontTag[i];
                               trimChars[i] = c;
                           }
                           sSource = sSource.TrimStart(trimChars);
                       }
                       */
                    if (sFrontTag == "'''''")
                    {
						//NewMessage.Show ("here1");
                        InlineItalic(0);
                        InlineBold(0);
                        
                    }
                    else
                        if (sFrontTag == "'''")
                        {
						//NewMessage.Show ("here");
                            InlineBold(0);
                            // oSelection.set_Style(ref controlFile.oBodyText);

                        }
                        else
                            if (sFrontTag == "''")
                            {
                                InlineItalic(0);
                            }
                            else

                                if (sFrontTag == "_")
                                {
                                    // November 2012
                                    // We only turn off udnerlining or italics if expressedly ordered to do so
                                    // that is, we no longer stop underlining unless we started underlining (this prevents us from ruining
                                    // a <past> section
                                    
                                        if (false == controlFile.UnderscoreKeep)
                                        {
                                            if (true == controlFile.UnderlineShouldBeItalicInstead)
                                            {
                                                InlineItalic(0);
                                            }
                                            else
                                                InlineUnderline(0);
                                        }
                                        else
                                        {
                                            // just leave the underscore alone
                                            InlineWrite("_");
                                        }
                                     
                                }
                                else
                                    if (sFrontTag == "---")
                                    {
                                        InlineStrikeThrough(0);
                                    }
                                    else

                                        if (sFrontTag == "<<<")
                                        {
                                            ChatMode(0);
                                        }

                                    else if (sFrontTag == "^")
                                    {
                                        InlineSuper(0);
                                    }
                                    else if (sFrontTag == "##")
                                    {
                                        InlineSub(0);
                                    }
                                    else

                                        if (sFrontTag == "|")
                                        {

                                            // add a | back if this is a link
                                            if (lastLink != "")
                                                sSource = "|" + sSource;
                                        }
                                        else if (sEndTag == "]]")
                                        {
					//	NewMessage.Show ("huh");
                                            // do nothing
                                            switch (controlFile.OptionalCode)
                                            {
                                                case 1: InlineBold(0); break;
                                                case 2: InlineItalic(0); break;
                                            }
                                        }

                    //     nFirstStart = sSource.IndexOf(sFrontTag);

                }
                else
                {
                    // if not end tag, we just break
                    break;
                }




            }
            string sLast = sSource;
            if ("" != sLast /*&& TEXT_NULL != sLast*/)
            {
                InlineWrite(sLast);
            }

            //TEXT_NULL is to prevent variable comment lines (empty variables) from adding linesapce) 

            if (true == AddLineFeed)
            {
                InlineWrite(Environment.NewLine);
            }






        }

        /// <summary>
        /// for text that indicates a text transcript of a chat
        // </summary>
        /// <param name="onoff">1 - means to activate it</param>
        protected virtual void  ChatMode(int onoff)
        {
            InlineWrite("***");
        }

        /// <summary>
        /// formats text, adding it to tallytable if necessary
        /// </summary>
        /// <param name="sBoldText"></param>
        /// <param name="secondtype">if true we don't use sBoldText as the code and we try to extract a new one</param>
        /// <returns></returns>
        protected string VariableCheck(string sBoldText, string sOriginal, string[] sFormats, bool secondtype)
        {
            //int start_index = sOriginal.IndexOf("[[");
            int nFirstBracket = sOriginal.IndexOf("[["); // in case on line with other brackets
            while (nFirstBracket > -1)
            {
                string sTitle = "";
               
                sTitle = sOriginal.Substring(0, nFirstBracket);




 

              //  [[~center]]'''[[~var title]]'''

                int end_index = -1;
                try
                {
                    end_index = sOriginal.IndexOf("]]");// Was in the wrong place,nFirstBracket); // in case on line with other brackets
                }
                catch (Exception ex)
                {
                    NewMessage.Show(ex.ToString());
                }

                string sCode = "";

                if (secondtype == true)
                {
                    sCode = sOriginal.Substring(nFirstBracket + 2, end_index - (nFirstBracket+2));
                }
                else
                {
                    sCode = sBoldText; // we take what was passed in (this is the normal way we parse special codes)
                }

                
                if ("" != sCode)
                {
                    // did not break existing tallies but need to extra [[f]] portion from here to store things correctly



                    foreach (string s in sFormats)
                    {
                        // strip all tagging from title
                        sTitle = sTitle.Replace(s, "");
                    }

                    // dec 2009
                    // - grab page number?
                    sTitle = String.Format("||{0}||{1}", sTitle, PageNumber());
                    sTitle = sTitle.Replace(",", " ");

                    // july 2010 - we add teh chapter
                    sTitle = "(Chapter "+ (chapter-1).ToString() + ") " + sTitle;

                    if (tallyTable.ContainsKey(sCode))
                    {
                        // add to string list
                        string sList = tallyTable[sCode].ToString();
                        sList = sList + ", " + sTitle;
                        tallyTable[sCode] = sList;
                    }
                    else
                    {
                        // start a tally 
                        tallyTable.Add(sCode, sTitle);

                    }


                    // we just strip the coding out of sBoldText?


              

                    sBoldText = sBoldText.Replace("[["+sCode+"]]", "").Trim();

                    // June 2011 - in case there are multiple codes in one line
                    // we don't want infinite loopage.
                    try
                    {
                        // June 6 2011 -- fixing a bug whereas we search outside the string
                        if ((nFirstBracket + 1) < sBoldText.Length)
                        {
                            nFirstBracket = sBoldText.IndexOf("[[", nFirstBracket + 1);
                        }
                        else
                        {
                            nFirstBracket = sBoldText.IndexOf("[[");
                        }
                    }
                    catch (Exception ex)
                    {
                        NewMessage.Show(ex.ToString());
                    }

                }
            }


            return sBoldText;

        }


        /// <summary>
        /// We lookup and ADD the footnote. sID will already be lower cased
        /// </summary>
        /// <param name="sID"></param>
        protected virtual void AddFootnote(string sID)
        {
        }


        /// <summary>
        /// For children like sendePub will store in a hash for use when generating ePub files
        /// </summary>
        /// <param name="id"></param>
        /// <param name="idtext"></param>
        protected virtual void AddId(string id, string idtext)
        {
        }

        /// <summary>
        /// get current page number
        /// </summary>
        /// <param name="nValue"></param>
        protected virtual string PageNumber()
        {
            return sLastTitle ;// "undefined page number";

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sTag">such as <b>bold me</b></param>
        /// <param name="sSource"></param>
        /// <param name="bMustBeAtStart">if true, which will generally be the case, will only returnt true if = starts the line</param>
        /// <returns>true if format specified</returns>
        private bool IsFormat(string sFrontTag, string sEndTag, ref string sSource, bool bMustBeAtStart)
        {

            if (bMustBeAtStart == true)
            {
                // if we do not start with the front tag, abort
                if (sSource.IndexOf(sFrontTag) != 0)
                {
                    return false;
                }
            }

            bool bFound = false;

            string exp = String.Format("(?<={0}).*(?={1})", sFrontTag, sEndTag);
            System.Text.RegularExpressions.Regex regex = new Regex(exp);

            //System.Text.RegularExpressions.MatchCOllection matches = FindSubstrings(sSource,exp, false);

            MatchCollection matches = regex.Matches(sSource);




            if (matches.Count > 0)
            {
                bFound = true;


                foreach (Match match in matches)
                {

                    sSource.Remove(match.Index, match.Length);
                    sSource.Insert(match.Index, match.Value);
                }




                sSource = sSource.Replace(sFrontTag, "");
                sSource = sSource.Replace(sEndTag, ""); // this will remove single =


                //  

            }


            return bFound;
        }

        /// <summary>
        /// Adds a new variable
        /// extracts from string of the format
        /// ~var=coin:Crown]]
        /// </summary>
        /// <param name="sAdd"></param>
        protected void VariableAdd(string sAdd)
        {
            sAdd = sAdd.Replace("]]", "");
            sAdd = sAdd.Replace("~var=", "");

            int nColon = sAdd.IndexOf(":");
            if (nColon > -1)
            {
                string sKey = sAdd.Substring(0, nColon).Trim();
                string sValue = sAdd.Substring(nColon+1, sAdd.Length - nColon-1).Trim();
               // System.Windows.Forms.MessageBox.Show(sKey + " " + sValue);
                variableTable.Add(sKey, sValue);
            }


        }
        /// <summary>
        /// We have a variable on this string, replace it with
        /// </summary>
        /// <param name="sSource"></param>
        protected string VariableSwap(string sSource)
        {
            string sValue = "";
            int nBrackets = sSource.Length - 1;// ("]]");
            if (nBrackets > -1)
            {
                // figure out the key we have
                
                // first space is the nex
                int nSpace = sSource.LastIndexOf(" "); // changed to last index of (June 2011) to 
                                                       //fix an issue with lines like[[~center]]'''[[~var title]]'''

                if (nSpace > nBrackets || nSpace == -1)
                {
                    throw new Exception("space comes after brackets (or not present). Not possible");
                }
                string sKey = sSource.Substring(nSpace, nBrackets - nSpace+1).Trim();


                sKey = sKey.Replace("]]", " ").Trim(); // june 2011 clearing when double variables on line

                if (variableTable.ContainsKey(sKey) == true)
                {
                    sValue = variableTable[sKey].ToString();
                }
                else
                {
                    sValue = "ERROR CANNOT FIND VARIABLE" + sKey;
                }
                //System.Windows.Forms.MessageBox.Show(sKey);
            }

            return sValue;
        }
        /// <summary>
        /// Is used to fix problems with numbered lists not restarting
        /// this is a list that is defined in config file of second/3rd level
        /// entries
        /// 
        /// i.e., if you are on tier 1 and find a i. that means 
        /// you need to reset the list
        /// </summary>
        /// <param name="sScan"></param>
        /// <returns></returns>
        protected  bool FixListScan(string sScan)
        {
            if (controlFile.FixNumberList != null)
            {
                string[] sListIndents = controlFile.FixNumberList;
                foreach (string s in sListIndents)
                {
                    if (s == sScan)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        ///////////////////////////////////
        //
        // OVERRIDE US
        //
        ///////////////////////////////////

        /// <summary>
        /// inserts a page break
        /// </summary>
        protected virtual void AddPageBreak()
        {
        }

        /// <summary>
        /// Align text until next alignment command is hit
        /// 0 - Center
        /// 1 - Left
        /// 2 - Right
        /// </summary>
        /// <param name="nAlignment"></param>
        protected virtual void AlignText(int nAlignment)
        {
        }

        /// <summary>
        /// Adds a bookmark at cursor point
        /// </summary>
        /// <param name="sBookmark"></param>
        protected virtual void AddBookmark(string sBookmark)
        {
        }

        /// <summary>
        /// adds a tab character or whatever at the location
        /// </summary>
        protected virtual void AddTab()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected virtual void AddTable(string sText)
        {
        }

        /// <summary>
        /// Adds a table of contents at location
        /// </summary>
        protected virtual void AddTableOfContents()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected virtual void AddNumberedBullets(string sText)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected virtual void AddBullets(string sText)
        {
        }

        /// <summary>
        /// the title has changed, what should we do? (In the case of sendePub, we start a new file)
        /// </summary>
        protected virtual void OnTitleChange()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected virtual string AddTitle(string sText)
        {
            // replace this with name of current Chapter
            string sOriginal = sText;
            // had to change the titling because Keeper was adding a title automatically with the [[title]] keyword
            // JUNE 2010 using both now.
            sText = sText.Replace("[[title_]]", chaptertoken);
            sText = sText.Replace("[[title]]", chaptertoken);

            // if we modified the text it means we hit a title which means we are onto a new chapter
            if (sText != sOriginal)
            {


                OnTitleChange();
               
                // increment current chapter
                chapter++;
            }
            return sText;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        /// <param name="nLevel"></param>
        protected virtual void AddHeader(string sText, int nLevel)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sFormatType"></param>
        /// <returns></returns>
        protected virtual bool AddStartMultiLineFormat(string sFormatType)
        {
            // starting a format change
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected virtual void AddEndMultiLineFormat()
        {
        }
      
        /// <summary>
        /// when word doc is finished do cleanpu here
        /// </summary>
        protected virtual void Cleanup()
        {
            // look for missing links -> bookmark combos
            foreach (string sLink in bookmarks)
            {
                if (sFulltext.IndexOf(String.Format("[Anchor {0}", sLink)) == -1)
                {
                    // anchor not present
                    UpdateErrors("Missing anchor for " + sLink);
                }
            }

            if (FootnoteLog != "")
            {
                string[] footnoteslinked = FootnoteLog.Split(new char[1] { ',' },StringSplitOptions.RemoveEmptyEntries);
                NewMessage.Show(String.Format("{0} Footnotes linked and they are {1}", footnoteslinked.Length, FootnoteLog));
            }

            if (controlFile.EndMessage != "")
            {
				if (false == SuppressMessages)
                NewMessage.Show(controlFile.EndMessage);
            }
        }


        /// <summary>
        /// THis value counts any stray [[ or ]] characters to alert that a
        /// formatting error may have occurred
        /// </summary>
        public int error_CountInlineBrackets = 0;

        /// <summary>
        /// While processing linline formatting (bold, et cetera), this is used to write out a line of text
        /// </summary>
        /// <param name="sText"></param>
        protected virtual void InlineWrite(string sText)
        {
            if (sText.IndexOf("[[") > -1 || sText.IndexOf("]]") > -1)
            {
                error_CountInlineBrackets++;
            }
        }

        protected virtual void InlineBold(int nValue)
        {
            
        }

        protected virtual void InlineItalic(int nValue)
        {
          
        }
        protected virtual void InlineStrikeThrough(int nValue)
        {
    
        }
        /// <summary>
        /// nvalue is ignored for underline
        /// </summary>
        /// <param name="nValue"></param>
        protected virtual void InlineUnderline(int nValue)
        {

        }

        /// <summary>
        /// nvalue is ignored for underline
        /// </summary>
        /// <param name="nValue"></param>
        protected virtual void InlineSuper(int nValue)
        {

        }
        /// <summary>
        /// nvalue is ignored for underline
        /// </summary>
        /// <param name="nValue"></param>
        protected virtual void InlineSub(int nValue)
        {
           
        }
        /// <summary>
        /// adds a picture
        /// </summary>
        /// <param name="nValue"></param>
        protected virtual void AddPicture(string sPathToFile)
        {
            if (File.Exists(sPathToFile) == false)
            {
                throw new Exception(sPathToFile + " does not exist!");
            }
           // System.Windows.Forms.MessageBox.Show(sPathToFile);
        }

        /// <summary>
        /// adds a link
        /// </summary>
        /// <param name="nValue"></param>
        protected virtual void AddLink(string sPathToFile, string sTitle)
        {
           
            // System.Windows.Forms.MessageBox.Show(sPathToFile);
        }

        /// <summary>
        /// adds a footnote to the has
        /// 
        /// working on multi line footnotes, using <br> to indicate them but each descendent needs to handle that in their own way
        /// </summary>
        /// <param name="sID"></param>
        /// <param name="sText"></param>
        private void AddFootnoteText(string sID, string sText)
        {
            sID = sID.ToLower();
            if (FootnoteHash.ContainsKey((object)sID) == true)
            {
                NewMessage.Show(String.Format("{0} already present in footnotes.", sID));
                return;
            }
            FootnoteHash.Add(sID, sText);

        }

        /// <summary>
        /// before whatever initial operations are required to open the filestream
        /// or whatever (in the case of Word Auto, will require global variables)
		/// error code of -1 means to abort the write process
        /// </summary>
        protected virtual int InitializeDocument(ControlFile _controlFile)
        {
            FootnoteHash = new Hashtable();
            Footnotes_in_specific_chapter = new Hashtable();
            controlFile = _controlFile;
            tallyTable = new System.Collections.Hashtable();
            variableTable = new System.Collections.Hashtable();

			return 1;
        }
		public override string ToString ()
		{
			return string.Format ("[sendBase: ]");
		}
    }
}
