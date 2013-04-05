// sendWord.cs
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


using Word = NetOffice.WordApi;
//using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using CoreUtilities;
/* Features
 * 
 *************
 * DONE
 *************
 * # = Bullets
 * * = bullets (can't get multi-level to work, and when they do, they'll only work if you have a style)
 * = Heading 1... Heading 4 =
 * '' Italic
 * ''' Bold
 * ||table||
 * 
 * 
 * 
 * 
 ************
 * TO DO
 ************
 * 
 */

namespace SendTextAway
{
    class sendWord : sendBase
    {

        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        object oFalse = false;
        object oTrue = true;
        Word._Application oWord;
        Word._Document oDoc;
        Word.Selection oSelection;
        object oStyle = null;
            object oHeader1 = null;
            object oHeader2 = null;
            object oHeader3 = null;
            object oHeader4 = null;
        object oHeader5 = null;
        object oBullet = null;
            object oBulletNumber = null;
            object oTableStyle = null;

            object oTitle = null;
        

        /// <summary>
        /// before whatever initial operations are required to open the filestream
        /// or whatever (in the case of Word Auto, will require global variables)
        /// </summary>
        protected override int InitializeDocument(ControlFile _controlFile)
        {
            //create a word object instead of globals, ie., WOrd(ControlFIle)
            base.InitializeDocument(_controlFile);
            

            oWord = new Word.Application();
            oWord.Visible = true;
            object template = controlFile.Template;
            oDoc = oWord.Documents.Add( template,  oMissing,
                 oMissing,  oMissing);

            oSelection = oWord.Selection;



            // load defaults from tempalte
            
            try
            {
                oHeader1 = controlFile.Heading1;
                oHeader2 = controlFile.Heading2; // write getmethods?
                oHeader3 = controlFile.Heading3;
                oHeader4 = controlFile.Heading4;
                oHeader5 = controlFile.Heading5;
                oBullet = controlFile.Bullet;
                oTitle = controlFile.ChapterTitle;
                oStyle = controlFile.BodyText;
                oBulletNumber = controlFile.BulletNumber;
                oTableStyle = controlFile.TableStyle;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
				return -1;
            }

            // test if style exists before being used
            bool bExists = false;
            foreach (Word.Style style in oDoc.Styles)
            {
                if (style.NameLocal == (oStyle).ToString())
                {
                    bExists = true;
                }
            }
            if (false == bExists)
            {
                NewMessage.Show(oStyle.ToString() + " did not exist");
            }

			oSelection.Style = oStyle;// (March 2013 - this might be reasonable replacement
            // set defaults
         //   oSelection.set_Style(ref oStyle); //this never seemed to happen 12/12/2010 (FIXED! The name was wrong, trying to come up with a better error check system
            Word.Style oSetStyle = (Word.Style)oSelection.Style;

           
            /* DId not work to test
            if (oSetStyle.NameLocal != (oStyle).ToString())
            {
                NewMessage.Show(oStyle.ToString() + " did not exist"); 
            }*/


			return 1;
        }

        /// <summary>
        /// overwrite the tab character being written
        /// </summary>
        protected override void AddTab()
        {
            InlineWrite("\t");
        }

        /// <summary>
        /// Builds a table row
        /// </summary>
        /// <param name="sText"></param>
        protected override void AddTable(string sText)
        {
		
			//March 2013 NetOffice.WordApi.Enums.WdDefaultTableBehavior.wdWord8TableBehavior;
            object objDefaultBehaviorWord8 = Word.Enums.WdDefaultTableBehavior.wdWord8TableBehavior;

            string[] Cols = sText.Split(new string[1] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            Word.Table oTable;

            oTable = oDoc.Tables.Add(oSelection.Range, 1, Cols.Length,  objDefaultBehaviorWord8,  oMissing);

            oTable.Style =  oTableStyle;
            // oTable.Range.ParagraphFormat.SpaceAfter = 10;
            // oSelection.SetRange(oTable.Range.End + 20, 1);
            //oTable.Borders = Word.WdBorderType.wdBorderBottom;


            for (int table = 0; table < Cols.Length; table++)
            {

              //  oTable.Cell(oTable.Rows.Count, table + 1).Range.Text = 
                Word.Range range = oTable.Cell(oTable.Rows.Count, table + 1).Range;
                oSelection.SetRange(range.Start, range.End);
                FormatRestOfText(Cols[table]);

            }

            // edit from the end

			// March 2013 Word.Range wrdRng =  oDoc.Bookmarks[oEndOfDoc].Range;
            Word.Range wrdRng = oDoc.Bookmarks[oEndOfDoc].Range;
            oSelection.SetRange(wrdRng.Start, wrdRng.End);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected override void AddNumberedBullets(string sText)
        {
            try
            {
                // Figured out how to have two chunks of text
                //if (oBulletNumber.ToString() != "bulletdefault")
                {

                    string sTabPrefix = "";
                    string sOldText = sText;

                    // add tabs

                    sText = sText.TrimStart('#').Trim();
                    int nInsertionPoint = oSelection.Start;
                    string sTextToAdd = sTabPrefix + sText;// +Environment.NewLine;

                    oSelection.Style =  oBulletNumber;
                    //  oSelection.set_Style(ref oBulletNumber);

                    object defaultlist = Word.Enums.WdDefaultListBehavior.wdWord10ListBehavior;


                    /*while (oSelection.Range.ListFormat.ListType != Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering)
                    {
                        oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);
                    }*/

                    //oSelection.Range.ListFormat.List.

                    oSelection.Range.ListFormat.ApplyNumberDefault( defaultlist);
                    oSelection.Range.ListFormat.ApplyNumberDefault( defaultlist);



                    //   oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);
                    //    oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);

                    /*  Word.ListFormat cleanFormat = null;

                      if (cleanFormat == null)
                      {
                         cleanFormat =  oSelection.Range.ListFormat;
                      }
                      oSelection.Range. = cleanFormat;
                      */
                    int nCount = 0; bool done = false;

                    //System.Windows.Forms.MessageBox.Show(oSelection.Range.ListFormat.ListValue.ToString());

                    oSelection.Range.ListFormat.ListLevelNumber = 1;



                    while (done == false)
                    {
                        // test to see if are an empty string
                        if (sOldText.Length > nCount + 1)
                        {
                            if (nCount > 0 && sOldText[nCount] == '#')
                            {

                                // tab blanked the text; 

                                /*
                                Word.Range oldrange = oSelection.Range;
                                oSelection.SetRange(nInsertionPoint, nInsertionPoint + sTextToAdd.Length);
                                oSelection.Range.InsertAlignmentTab((int)Word.WdAlignmentTabAlignment.wdRight, (int)Word.WdAlignmentTabAlignment.wdLeft);
                                oSelection.SetRange(oldrange.Start, oldrange.End);
                                */

                                oSelection.Range.ListFormat.ListLevelNumber++; // ++crash
                                // oSelection.Range.ListFormat.ListIndent(); not what we want
                                //oSelection.TypeText("\t"); // closet but ick
                                //  oSelection.TypeText("\xB"); vertical tab did not work
                                // oSelection.TypeText();
                            }
                            else if (nCount > 0)
                            {
                                done = true;
                            }
                        }
                        else
                        {
                            done = true;
                        }
                        nCount++;
                    }

                    // we restart a list if the ListValue is greater than ONE 
                    // and we have reset InAList = false
                    // This means that we are at #2 of a list that thinks it is continuing
                    // from another list seperated by text

                    bool bRestartList = false;

                    if (oSelection.Range.ListFormat.ListValue > 1 && InAList == false)
                    {
                        bRestartList = true;
                    }



                    // we sometimes reset the list due to 'issues'
                    if ((oSelection.Range.ListFormat.ListLevelNumber == 1
                        && FixListScan(oSelection.Range.ListFormat.ListString) == true)
                        || (bRestartList == true)
                        )
                    {
                        foreach (Word.Paragraph para in oSelection.Range.ListParagraphs)
                        {

                            // para.ResetAdvanceTo();
                            //   para.SeparateList();
                            para.Reset(); // worked for simple list but not big doc
                            //para.Reset();
                        }
                    }

                    // set that we are in a list
                    //   if (oSelection.Range.ListFormat.ListValue == 1)
                    {
                        InAList = true;
                    }
                    /* worked but Reset seems easier
                     * object oDef = Word.WdDefaultListBehavior.wdWord10ListBehavior;
                     object myTemplateIndex = 1;

                     Word.ListTemplate myTemplate =
                         oWord.ListGalleries[Microsoft.Office.Interop.Word.WdListGalleryType.wdNumberGallery].ListTemplates.get_Item(ref myTemplateIndex);
                     oSelection.Range.ListFormat.ApplyListTemplate(myTemplate, ref oFalse, ref oFalse,
                        ref oDef );
                     * */

                    // System.Windows.Forms.MessageBox.Show("bad list");
                    /*   foreach (Word.Paragraph para in oSelection.Range.ListParagraphs)
                    
                       {
                       
                          // para.ResetAdvanceTo();
                        //   para.SeparateList();
                           para.Reset(); // worked for simple list but not big doc
                           para.Reset();
                       }
                       */
                    //  oSelection.Range.ListParagraphs[1].SeparateList();
                    //oSelection.Range.ListFormat.ApplyListTemplate(oSelection.Range.ListParagraphs[1].

                    // oSelection.Range.SetListLevel(1);

                    // oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);
                    // oSelection.set_Style(ref oBulletNumber);
                    // oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);


                    //oSelection.TypeText(sTextToAdd);
                    FormatRestOfText(sTextToAdd);// + "(" + oSelection.Range.ListFormat.ListValue.ToString());







                    //indent for each #

                    // turn off
                    //   oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);

                    // oSelection.Range.ListFormat.ListLevelNumber = 1;
                    oSelection.Range.ListFormat.ApplyNumberDefault( defaultlist);
                    oSelection.Style =oStyle;
                }
            }
            catch (Exception ex)
            {
                UpdateErrors("Failure in AddNumberedBullets + " + sText + " " + ex.ToString());
            }
            
            /*  else
              {
                  object defaultlist = Word.WdDefaultListBehavior.wdWord10ListBehavior;
                  oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);

                            
                  oSelection.TypeText(sText + Environment.NewLine);
                 // oSelection.Select(); did I need this

                  oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);
              }*/
            /* bool done = false;
             int nCount = 0;
             //indent for each #
             while (done == false)
             {
                 if (sText[nCount] == '#')
                 {
                     oSelection.Range.ListFormat.ListIndent();
                 }
                 else
                 {
                     done = true;
                 }
                 nCount++;
             }

                      
             sText = sText.TrimStart('#').Trim();



             object defaultlist = Word.WdDefaultListBehavior.wdWord10ListBehavior;
             oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);
                        
             oSelection.TypeText(sText + Environment.NewLine);

             oSelection.Range.ListFormat.ApplyNumberDefault(ref defaultlist);*/
        }

        /// <summary>
        /// Align text until next alignment command is hit
        /// 0 - Center
        /// 1 - Left
        /// 2 - Right
        /// </summary>
        /// <param name="nAlignment"></param>
        protected override void AlignText(int nAlignment)
        {
            switch (nAlignment)
            {
                    
                case 0: oSelection.ParagraphFormat.Alignment = Word.Enums.WdParagraphAlignment.wdAlignParagraphCenter; break;
			case 1: oSelection.ParagraphFormat.Alignment = Word.Enums.WdParagraphAlignment.wdAlignParagraphLeft; break;
			case 2: oSelection.ParagraphFormat.Alignment = Word.Enums.WdParagraphAlignment.wdAlignParagraphRight; break;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected override void AddBullets(string sText)
        {
          /*  sText = sText.TrimStart('*').Trim();

            object defaultlist = Word.WdDefaultListBehavior.wdWord10ListBehavior;
            oSelection.Range.ListFormat.ApplyBulletDefault(ref defaultlist);
            oSelection.Range.ListFormat.ListOutdent();
            //oSelection.TypeText(FormatRestOfText(sText) + Environment.NewLine);
          //  oSelection.Range.ListFormat.ListLevelNumber++;

            FormatRestOfText(sText);
            oSelection.Range.ListFormat.ApplyBulletDefault(ref defaultlist);*/



                string sOldText = sText;

                // add tabs

                sText = sText.TrimStart('*').Trim();
                int nInsertionPoint = oSelection.Start;
                string sTextToAdd = sText;

                oSelection.Style  =  oBullet;
            
                object defaultlist = Word.Enums.WdDefaultListBehavior.wdWord10ListBehavior;
                                             
                oSelection.Range.ListFormat.ApplyBulletDefault( defaultlist);
                oSelection.Range.ListFormat.ApplyBulletDefault( defaultlist);

                int nCount = 0; bool done = false;

            oSelection.Range.ListFormat.ListLevelNumber = 1;
                while (done == false)
                {
                    // test to see if are an empty string
                    if (sOldText.Length > nCount + 1)
                    {
                        if (nCount > 0 && sOldText[nCount] == '*')
                        {


                            oSelection.Range.ListFormat.ListLevelNumber++;
                        }
                        else if (nCount > 0)
                        {
                            done = true;
                        }
                    }
                    else
                    {
                        done = true;
                    }
                    nCount++;
                }

              
                FormatRestOfText(sTextToAdd);// + "(" + oSelection.Range.ListFormat.ListValue.ToString());


                oSelection.Range.ListFormat.ApplyBulletDefault( defaultlist);
                oSelection.Style = oStyle;
            
          
        }

        /// <summary>
        /// insert a page break, but only if we need it
        /// </summary>
        protected override void AddPageBreak()
        {
            Word.Enums.WdInformation info =Word.Enums.WdInformation.wdFirstCharacterLineNumber;
            int line = (int)oSelection.get_Information(info);
            int col =(int)oSelection.get_Information(Word.Enums.WdInformation.wdFirstCharacterColumnNumber);
            // if we are not on the first line and column of a page we insert a page break
            if ( !(line == 1 && col == 1) )
            {
                object pageBreak = Word.Enums.WdBreakType.wdPageBreak;
                oSelection.InsertBreak( pageBreak);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        protected override string AddTitle(string sText)
        {
            
            sText = base.AddTitle(sText);

            if (chapter > 1)
            {
                AddPageBreak();
            }

            


            oSelection.Style = oTitle;
            oSelection.TypeText(sText + Environment.NewLine);
            oSelection.Style = oStyle;
            return sText;
        }
        /// <summary>
        /// sID is already lowercase
        /// </summary>
        /// <param name="sID"></param>
        protected override void AddFootnote(string sID)
        {
            sID = sID.Trim();
            if (FootnoteHash.ContainsKey(sID) == true)
            {
                string sSource = (string) FootnoteHash[sID];

                sSource = sSource.Replace("<br>", Environment.NewLine);
                // process text lienfeed

                object sText = sSource;
                object id = (object)sID;

              //  oSelection.Text = oSelection.Text + " ";
                oSelection.Select();
                object start = oSelection.Range.End-2;
                
                object end = oSelection.Range.End-1;// oSelection.Range.End - 1;
                Word.Range mark = oDoc.Range( start,  end);
                //oSelection.TypeText("(Note)");

               

                
                oSelection.Footnotes.Add(mark,  oMissing,  sText);
                object amount = 1;
                oSelection.MoveLeft( oMissing,  amount,  oMissing);
            }
            else
            {
                NewMessage.Show(String.Format("{0} footnote not found!", sID));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sText"></param>
        /// <param name="nLevel">Heading level 1..4</param>
        protected override void AddHeader(string sText, int nLevel)
        {
            object oStyleToUse = null;
            switch (nLevel)
            {
                case 1: oStyleToUse = oHeader1; break;
                case 2: oStyleToUse = oHeader2; break;
                case 3: oStyleToUse = oHeader3; break;
                case 4: oStyleToUse = oHeader4; break;
                case 5: oStyleToUse = oHeader5; break;
            }
            oSelection.Style =  oStyleToUse;
            oSelection.TypeText(sText + Environment.NewLine);
            oSelection.Style = oStyle;
        }

        /// <summary>
        /// Starts a multiline format change such as
        /// <example>
        /// here
        /// </example>
        /// </summary>
        /// <param name="sFormatType"></param>
        /// <returns></returns>
        protected override bool AddStartMultiLineFormat(string sFormatType)
        {
              // starting a format change
            object newFormat = sFormatType;
            oSelection.Style = newFormat;
           return true;
        }

        /// <summary>
        /// resets the formatting of a multi line format to normal format
        /// </summary>
        /// <returns></returns>
        protected override void AddEndMultiLineFormat()
        {
            oSelection.Style = oStyle;
        }

        protected override void ChatMode(int onoff)
        {
            switch (controlFile.ChatMode)
            {
                case 0: InlineUnderline(onoff); break;
                case 1: InlineItalic(onoff); break;
                case 2: InlineBold(onoff); break;
            }
          
       /*     if (1 == onoff)
            {
               // object newFormat = "Emphasis";
               // oSelection.set_Style(ref newFormat);
                InlineUnderline(1);
            }
            else
            {
                // back to default
               // oSelection.set_Style(ref oStyle);
            }*/
        }
        protected override int LineCount()
        {
          //  return oDoc.Sentences.Count;
            return oDoc.Paragraphs.Count;
            
        }

        /// <summary>
        /// searches for the object in text
        /// 
        /// May 2012
        /// Original intent was just to look for extra brackets in text as
        /// error detction mechanism
        /// </summary>
        /// <param name="searchMe"></param>
        protected override void SearchFor(object searchMe)
        {
         /* NEVER WORKED    
            object item = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine;
            object whichItem = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst;
            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
            object forward = false; //cause we are back of it?
            object matchAllWord = false;
            object missing = System.Reflection.Missing.Value;
 
            oSelection.Range.Document.GoTo(ref item, ref whichItem, ref missing, ref missing);
 
            oSelection.Range.Find.Execute(ref searchMe, ref missing, ref matchAllWord,
               ref missing, ref missing, ref missing, ref forward,
               ref missing, ref missing, ref missing, ref replaceAll,
               ref missing, ref missing, ref missing, ref missing);
        */
        }
    
   
        protected override void Cleanup()
        {
            /*12/12/2010 - disabled, lost text between page breaks
            if (controlFile.Linespace != -1)
            {
                // do cleanup
                // fix line spacing
                NewMessage.Show("Starting Double Space -- DO NOT TOUCH WORD FILE UNTIL DONE");
                foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in oDoc.Paragraphs)
                {
                    paragraph.Format.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
                    paragraph.Format.LineSpacing = oWord.LinesToPoints(controlFile.Linespace);
                }
                NewMessage.Show("DOUBLE SPACE FINISHED");
            }*/

            //* trying to replace tabbed underlines with no underline... instead created a macro and a reminder to run it
            /*
            oSelection.Find.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
            oSelection.Find.Replacement.ClearFormatting();
            oSelection.Find.Replacement.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
            oSelection.Find.Text = "^t";
            oSelection.Find.Replacement.Text = "^t";
            oSelection.Find.
            */

            // may 2012 trying to count lines
          /*  foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in oDoc.Paragraphs)
            { 
             // count paragraphs
                foreach (
            }*/
           


            base.Cleanup();

            // update the table of contents if present
            if (oDoc.TablesOfContents.Count > 0)
            {
                oDoc.TablesOfContents[1].Update();
            }

          

            

        }

        
        /// <summary>
        /// While processing linline formatting (bold, et cetera), this is used to write out a line of text
        /// </summary>
        /// <param name="sText"></param>
        protected override void InlineWrite(string sText)
        {
            base.InlineWrite(sText);
            oSelection.TypeText(sText);
        }

        protected override void InlineBold(int nValue)
        {
            oSelection.Font.Bold = nValue;
        }

        protected override void InlineItalic(int nValue)
        {
            oSelection.Font.Italic = nValue;
            
        }

      /*  /// <summary>
        /// get current page number
        /// </summary>
        /// <param name="nValue"></param>
        protected override string PageNumber()
        {
           // return oSelection.Footnotes[0].ToString();
            
            return oSelection.Footnotes.Count.ToString();

        }*/
        protected override void InlineStrikeThrough(int nValue)
        {
            oSelection.Font.StrikeThrough = nValue;
        }
        /// <summary>
        /// nvalue is ignored for underline
        /// </summary>
        /// <param name="nValue"></param>
        protected override void InlineSuper(int nValue)
        {
            oSelection.Font.Superscript = nValue;
        }

        /// <summary>
        /// nvalue is ignored for underline
        /// </summary>
        /// <param name="nValue"></param>
        protected override void InlineSub(int nValue)
        {
            oSelection.Font.Subscript = nValue;
        }
        /// <summary>
        /// Adds a bookmark at cursor point
        /// </summary>
        /// <param name="sBookmark"></param>
        protected override void AddBookmark(string sBookmark)
        {
            sBookmark = sBookmark.Trim();
            object Range = oSelection.Range;
            oSelection.Bookmarks.Add(sBookmark,  Range);

        

        }
        int formatCount = 0; // Nov 2012 will increment. If we are already underlining and then we have underlining inside of underlining
                             // we do not STOP underlining until the block text is finished formatting.
        /// <summary>
        /// nvalue is ignored for underline
        /// </summary>
        /// <param name="nValue"></param>
        protected override void InlineUnderline(int nValue)
        {



            if (nValue == 1)
            {
                if (oSelection.Font.Underline == Word.Enums.WdUnderline.wdUnderlineSingle)
                {
                    NewMessage.Show("We were asked to start underining during a segment of text wherein we were ALREADY underlining. We are ignoring this request. Check for any text issues.");
                    formatCount++;
                }
                oSelection.Font.Underline = Word.Enums.WdUnderline.wdUnderlineSingle;
            }
            else
                if (nValue == 0)
                {
                    if (0 == formatCount)
                    {
                        oSelection.Font.Underline = Word.Enums.WdUnderline.wdUnderlineNone;
                    }
                    else formatCount--;
                    

                }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sPathToFile"></param>
        protected override void AddLink(string sPathToFile, string sTitle)
        {
            object oTextToShow = sTitle;
            object oLink = sPathToFile;
            oSelection.Hyperlinks.Add(oSelection.Range,  oLink,   oMissing,
                 oMissing,  oTextToShow,  oMissing);
            // System.Windows.Forms.MessageBox.Show(sPathToFile);
        }
        /// <summary>
        /// adds a picture
        /// </summary>
        /// <param name="nValue"></param>
        protected override void AddPicture(string sPathToFile)
        {
            base.AddPicture(sPathToFile);

            object oWidth = 10;
            object oLeft = 1;
            object oTop = oSelection.Range.Start;
            object oHeight = 100;
            // System.Windows.Forms.MessageBox.Show(sPathToFile);
            //oDoc.Shapes.AddPicture(sPathToFile, ref oMissing, ref oMissing, ref oLeft, ref oTop,                ref oWidth, ref oHeight, ref oMissing);
            oSelection.InlineShapes.AddPicture(sPathToFile,  oMissing,  oTrue,  oMissing);
        }

        /// <summary>
        /// Adds a table of contents at location
        /// </summary>
        protected override void AddTableOfContents()
        {
            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "3";
            Object oTrue = true;
            Object oTOCTableID = "TableOfContents";
            
            
            oDoc.TablesOfContents.Add(oSelection.Range,  oTrue,  oUpperHeadingLevel,
                oLowerHeadingLevel,  oMissing,  oTOCTableID,  oTrue,
                oTrue,  oMissing,  oTrue,  oTrue,  oTrue);



        }

        /// <summary>
        /// just an example function
        /// </summary>
        public void example()
        {

       

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add( oMissing,  oMissing,
                 oMissing,  oMissing);

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add( oMissing);
            object styleHeading1 = "Heading 1";
            oPara1.Style =  styleHeading1;
            oPara1.Range.Text = "Heading 1";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks[ oEndOfDoc].Range;
            oPara2 = oDoc.Content.Paragraphs.Add( oRng);
            oPara2.Range.Text = "Heading 2";
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();

            //Insert another paragraph.
            Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks[ oEndOfDoc].Range;
            oPara3 = oDoc.Content.Paragraphs.Add( oRng);
            oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 24;
            oPara3.Range.InsertParagraphAfter();

            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks[ oEndOfDoc].Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 5,  oMissing,  oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;
            string strText;
            for (r = 1; r <= 3; r++)
                for (c = 1; c <= 5; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;

            //Add some text after the table.
            Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks[ oEndOfDoc].Range;
            oPara4 = oDoc.Content.Paragraphs.Add( oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "And here's another table:";
            oPara4.Format.SpaceAfter = 24;
            oPara4.Range.InsertParagraphAfter();

            //Insert a 5 x 2 table, fill it with data, and change the column widths.
            wrdRng = oDoc.Bookmarks[ oEndOfDoc].Range;
            oTable = oDoc.Tables.Add(wrdRng, 5, 2,  oMissing,  oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            for (r = 1; r <= 5; r++)
                for (c = 1; c <= 2; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            oTable.Columns[2].Width = oWord.InchesToPoints(3);

            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            object oPos;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks[ oEndOfDoc].Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks[ oEndOfDoc].Range;
                wrdRng.ParagraphFormat.SpaceAfter = 6;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                               (Word.Enums.WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));
            object oCollapseEnd = Word.Enums.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.Enums.WdBreakType.wdPageBreak;
            wrdRng.Collapse( oCollapseEnd);
            wrdRng.InsertBreak( oPageBreak);
            wrdRng.Collapse( oCollapseEnd);
            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();

            //Insert a chart.
//            Word.InlineShape oShape;
//            object oClassType = "MSGraph.Chart.8";
//            wrdRng = oDoc.Bookmarks[ oEndOfDoc].Range;
//            oShape = wrdRng.InlineShapes.AddOLEObject( oClassType,  oMissing,
//                 oMissing,  oMissing,  oMissing,
//                 oMissing,  oMissing,  oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
//            object oChart;
//            object oChartApp;
//            oChart = oShape.OLEFormat.Object;
//            oChartApp = oChart.GetType().InvokeMember("Application",
//                BindingFlags.GetProperty, null, oChart, null);
//
//            //Change the chart type to Line.
//            object[] Parameters = new Object[1];
//            Parameters[0] = 4; //xlLine = 4
//            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
//                null, oChart, Parameters);
//
//            //Update the chart image and quit MSGraph.
//            oChartApp.GetType().InvokeMember("Update",
//                BindingFlags.InvokeMethod, null, oChartApp, null);
//            oChartApp.GetType().InvokeMember("Quit",
//                BindingFlags.InvokeMethod, null, oChartApp, null);
//            //... If desired, you can proceed from here using the Microsoft Graph 
//            //Object model on the oChart and oChartApp objects to make additional
//            //changes to the chart.
//
//            //Set the width of the chart.
//            oShape.Width = oWord.InchesToPoints(6.25f);
//            oShape.Height = oWord.InchesToPoints(3.57f);

            //Add text after the chart.
            wrdRng = oDoc.Bookmarks[ oEndOfDoc].Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("THE END.");

            //Close this form.
            


        }
        /// <summary>
        /// This did not work
        /// </summary>
        /// <param name="source"></param>
        /// <param name="matchPattern"></param>
        /// <param name="findAllUnique"></param>
        /// <returns></returns>
        public static Match[] FindSubstrings(string source,
   string matchPattern, bool findAllUnique)
        {
            SortedList uniqueMatches = new SortedList();
            Match[] retArray = null;
            Regex RE = new Regex(matchPattern, RegexOptions.Multiline);
            MatchCollection theMatches = RE.Matches(source);
            if (findAllUnique)
            {
                for (int counter = 0; counter < theMatches.Count; counter++)
                {
                    if (!uniqueMatches.ContainsKey(theMatches[counter].Value))
                    {
                        uniqueMatches.Add(theMatches[counter].Value,
                        theMatches[counter]);
                    }
                }
                retArray = new Match[uniqueMatches.Count];
                uniqueMatches.Values.CopyTo(retArray, 0);
            }
            else
            {
                retArray = new Match[theMatches.Count];
                theMatches.CopyTo(retArray, 0);
            }
            return (retArray);
        }
		public override string ToString ()
		{
			return string.Format ("[sendWord]");
		}
    }
}


