// sendePub2.cs
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
using CoreUtilities;
using System.Collections;

// note: I started trying to fix this but instead created a new class to correct issues
namespace SendTextAway
{
	/// <summary>
	/// This will eventually create a full epub folder.
	/// 
	
	/// </summary>
	public class sendePub2 : sendPlainText
	{
		string ZIP7 = "";//@"C:\Program Files\7-Zip\7z.exe";
		string directory_to_sourcefiles = "";// @"C:\Users\Brent\Documents\Visual Studio 2005\Projects\SendTextAway\SendTextAway\bin\Debug\epubfiles\";

		private int NeedParagraphClosingCounter = 0;
		//	private bool NeedParagraphClosing; // used for centering. If true the next InlineWrite will write a </p> closing tag
		private string sDirectory = "";
		ArrayList footnotesincurrentsection; // footnotes can be chapter, or if I can figure out a good solution the entire document
		
		private Hashtable ids;
		
		public static string GetDateDirectory {
			get { return DateTime.Today.ToString ("yyyy MM dd");}
		}

		protected override string HandleEmDash (string sText)
		{
			// we replace any existing ampersands with the code before applying Fancy

			// we do not want to replace any of the FancyCharacter codes
			// so we ensure no semi-colon.
			// (but realistically we should not need this...)
			//sText = sText.Replace ("&", "&amp;");
			if (controlFile.Emdash_removespacesbeforeafter == true) {
				// removing spaces
				sText = sText.Replace (" -- ", "&#8212;");
				sText = sText.Replace (" --", "&#8212;");

				sText = sText.Replace ("-- ", "&#8212;");
			}
			return sText.Replace ("--", "&#8212;");
		}

		protected override string FixHTMLEncoding (string sText)
																																															{

// 21/05/2014
// I want to be able to output stuff like &quot and &lt to escape characters -- so I'm thinking of just doing an exclusion test here. I'm sure there's better ways...
// theoretically if I 'escaped' all the text correctly via a manual means -- as I have done in this case -- all should be fine
if (sText.IndexOf ("&quot;") <= -1 && sText.IndexOf ("&lt;") <= -1 && sText.IndexOf ("&gt;") <= -1) {
sText = sText.Replace ("&", "&amp;");
					}
			sText = sText.Replace ("…", "...");
			// Unicode generic replacement http://stackoverflow.com/questions/1488866/how-to-replace-i-in-a-string
			//sText = sText.Replace ("\xEF\xBF\xBD","...");
			sText = sText.Replace ("\xFFFD", "...");
			return sText;
		}

		protected override string ReplaceFancyCharacters (string sSource)
		{
	

			if (true == controlFile.FancyCharacters) {
				
				sSource = sSource.Replace (".\"", ".&#8221;");
				
				sSource = sSource.Replace ("\".", "&#8221;.");
				
				
				sSource = sSource.Replace ("!\"", "!&#8221;");
				sSource = sSource.Replace ("?\"", "?&#8221;");
				sSource = sSource.Replace ("--\"", "--&#8221;");
				
				sSource = sSource.Replace ("-\"", "-&#8221;");
				
				sSource = sSource.Replace ("-- \"", "-- &#8221;");
				sSource = sSource.Replace (",\"", ",&#8221;");
				// remainder of quotes will be the opposite way
				sSource = sSource.Replace ("\"", "&#8220;");
				
				// do standard repalcements (February 2010)
				if (sSource.IndexOf ("...") > -1) {
					sSource = sSource.Replace ("...", "&hellip;");
				}
				
				// to finish look here: http://www.unicode.org/charts/charindex.html
			}
			return sSource;
		}
		/// <summary>
		/// before whatever initial operations are required to open the filestream
		/// or whatever (in the case of Word Auto, will require global variables)
		/// </summary>
		protected override int InitializeDocument (ControlFile _controlFile)
		{
			directory_to_sourcefiles = _controlFile.TemplateDirectory;
			ZIP7 = _controlFile.Zipper;
			
			if (null == directory_to_sourcefiles || Constants.BLANK == directory_to_sourcefiles) {
				NewMessage.Show (Loc.Instance.GetString ("To generate an epub you need to specify a valid path to the template files used to generate the final epub files"));
				return -1;
			}
			if (null == ZIP7 || Constants.BLANK == ZIP7) {
				NewMessage.Show (Loc.Instance.GetString ("No path to 7-Zip was specified. This step will be skipped.)"));
			}
			// create an output directory based on date
			sDirectory = Path.Combine (_controlFile.OutputDirectory, GetDateDirectory);
			Directory.CreateDirectory (sDirectory);
			File.Copy (Path.Combine (directory_to_sourcefiles, "mimetype"), Path.Combine (sDirectory, "mimetype"), true);
			
			Directory.CreateDirectory (Path.Combine (sDirectory, "oebps"));
			DirectoryInfo images = Directory.CreateDirectory (Path.Combine (sDirectory, "oebps\\images"));
			Directory.CreateDirectory (Path.Combine (sDirectory, "META-INF"));
			
			// start with footer as default first file?
			currentFileBeingWritten = Path.Combine (sDirectory, "oebps\\" + "preface.xhtml");
			
			chapter = _controlFile.StartingChapter;
			base.InitializeDocument (_controlFile);
			
			
			// copy required files
			
			// copy image files
			FileUtils.Copy (new DirectoryInfo (Path.Combine (directory_to_sourcefiles, "oebps\\images")), images, "*.*", "*.*", false, 2, new System.Windows.Forms.ProgressBar ());
			
			if (true == _controlFile.CopyTitleAndLegalTemplates) {
				File.Copy (Path.Combine (directory_to_sourcefiles, "oebps\\copyright.xhtml"), Path.Combine (sDirectory, "oebps\\copyright.xhtml"), true);
				File.Copy (Path.Combine (directory_to_sourcefiles, "oebps\\legal.xhtml"), Path.Combine (sDirectory, "oebps\\legal.xhtml"), true);
				File.Copy (Path.Combine (directory_to_sourcefiles, "oebps\\title_page.xhtml"), Path.Combine (sDirectory, "oebps\\title_page.xhtml"), true);
			}


			if (File.Exists (controlFile.OverrideStyleSheet)) {
				File.Copy (controlFile.OverrideStyleSheet, Path.Combine (sDirectory, "oebps\\stylesheet.css"), true);
			} else {
				// copy the stock file
				File.Copy (Path.Combine (directory_to_sourcefiles, "oebps\\stylesheet.css"), Path.Combine (sDirectory, "oebps\\stylesheet.css"), true);
			}
			File.Copy (Path.Combine (directory_to_sourcefiles, "oebps\\page-template.xpgt"), Path.Combine (sDirectory, "oebps\\page-template.xpgt"), true);
			
			
			File.Copy (Path.Combine (directory_to_sourcefiles, "META-INF\\container.xml"), Path.Combine (sDirectory, "META-INF\\container.xml"), true);
			footnotesincurrentsection = new ArrayList ();
			FootnoteChapterIAmInHash = new Hashtable ();
			ids = new Hashtable ();
			
			
			return 1;
		}
		
		
		protected override void ChatMode (int onoff)
		{
			switch (controlFile.ChatMode) {
			case 0: if (onoff == 1) {
					InlineWrite ("<u>");
				} else {
					InlineWrite ("</u>");
				}
				break;
			case 1: if (onoff == 1) {
					InlineWrite ("<i>");
				} else {
					InlineWrite ("</i>");
				}
				break;
			case 2: if (onoff == 1) {
					InlineWrite ("<b>");
				} else {
					InlineWrite ("</b>");
				}
				break;
			}
		}
		/// <summary>
		/// generic Insert function for InsertHeader and InsertFooter
		/// </summary>
		/// <param name="sTemplate"></param>
		/// <param name="sChapterToken"></param>
		private void InsertFromTemplate (string sTemplate, string sChapterToken)
		{
			if (null == file1) {
				throw new Exception ("no file open for writing");
			}
			// loads a header file
			StreamReader header = new StreamReader (sTemplate);
			if (null != header) {
				string line = header.ReadLine ();
				while (line != null) {
					if ("" != sChapterToken) {
						line = line.Replace ("[chapter]", sChapterToken);
						
					}
					file1.WriteLine (line);
					line = header.ReadLine ();
				}
			}
			header.Close ();
			header.Dispose ();
		}
		/// <summary>
		/// inserts text for header file
		/// <param name="sChapterToken">If not "" then if we find [chapter] we replace it with sChapterToken</param>
		/// </summary>
		protected override void InsertHeaderForNewFile (string sChapterToken)
		{
			InsertFromTemplate (Path.Combine (directory_to_sourcefiles, "header.txt"), sChapterToken);
			
		}
		
		protected override void InsertFooterorNewFile (string sChapterToken)
		{
			InsertFromTemplate (Path.Combine (directory_to_sourcefiles, "footer.txt"), sChapterToken);
		}
		
		/// <summary>
		/// current chapter file being closed off
		/// </summary>
		protected override void CloseCurrentFile ()
		{
			

			
			// blank list of footnotes for this current section (MIGHT BE MOVED IF CHANGE FOOTNOTE SYSTEM)
			// June 2010 now writign to end file instead footnotesincurrentsection = new ArrayList(); ;
			
			
		
			base.CloseCurrentFile ();
			
			
			
			
			
		}

		private int TabParagraph = 0; // count how many <p> tagsa we've introduce through plain, unformatted text
		/// <summary>
		/// While processing linline formatting (bold, et cetera), this is used to write out a line of text
		/// </summary>
		/// <param name="sText"></param>
		protected override void InlineWrite (string sText)
		{
			bool debugOnlyIsBlank = false;
			if (sText == Constants.BLANK) {
				debugOnlyIsBlank = true;
				//NewMessage.Show ("sText is blank");
			}

			if (true == adding_table && false == adding_row_to_table) {
				adding_table = false;
				file1.WriteLine ("</table>");
			}
			
			
			// if we were writing a bullet list, close it
			// this adds the </ul>
			if ("" != startnormalbullet && false == adding_bullet_now) {
				file1.Write (startnormalbullet);
				startnormalbullet = "";
			}



			sText = sText.Replace (Environment.NewLine, "<p></p>");
		

			//	 "{0}";



			// we only do a <p> break if we detect a tab (May 2013 - - looks like I removed this
			if (sText.IndexOf ("\t") > -1) {

				
				//NeedParagraphClosing = true;
				
			}

			// replacing NeedParagraphClosing with a counter


			
//			if (true == NeedParagraphClosing && "" != sText )
//			{
//				sLine = "{0}</p>";
//				NeedParagraphClosing = false;
//			}

			string paragraphclosers = "";
			// 20/04/2014 - do not add closers IF we are in midst of formating like (i.e., <p><u>hi</u></p> NOT <p><u></p></u>
			// May 2013 - do not add closers unless they are closing text
			if (sText != Constants.BLANK && sText != " " && IsFormating == false) {
				while (NeedParagraphClosingCounter > 0) {
				//	NewMessage.Show ("for==" + sText);
					paragraphclosers = paragraphclosers + "</p>";
					NeedParagraphClosingCounter--;
				}
			}
			string sLine = sText + paragraphclosers;
			;
			//sLine = String.Format(sLine, sText);

			/* I do not know what this is For Feb 4 2014
			 * I do not understand what the plan here was. This broke HTMl ecoding.*/
			sLine = sLine.Replace ("[<", "&lt;");
			sLine = sLine.Replace ("[>", "&gt;");


			if (debugOnlyIsBlank) {
			//	NewMessage.Show ("We had a blank line. Now we decide to write out: "+sLine);
			}
			//else // If you start off blank do you ever want to write anything 19/06/2014 -- This might be a very bad idea. Trying to resolve Novel_Facts_Paragraphs text. DID NOT FIX ANYTHING SO I RESET THIS
			file1.Write (sLine);


			DoErrorTest (sLine);
			//oSelection.TypeText(sText);
		}
//		private bool IsFormating (string sTag)
//		{
//			if (sTag == "<u>" || sTag =="</u>") {
//				return true;
//			}
//
//			return false;
//		}
		private string lastmultiline = "";

		protected override bool AddStartMultiLineFormat (string sFormatType)
		{
			InlineWrite ("<div class=\"" + sFormatType + "\">");
			lastmultiline = sFormatType;
			return true;
		}
		
		/// <summary>
		/// resets the formatting of a multi line format to normal format
		/// </summary>
		/// <returns></returns>
		protected override void AddEndMultiLineFormat ()
		{
			if ("" != lastmultiline) {
				//InlineWrite("</"+lastmultiline+">");
				InlineWrite ("</div>");
			}
		}
		
		bool adding_table;
		bool adding_row_to_table; // while adding row we don't put a /table
		protected override void AddTable (string sText)
		{
			if (false == adding_table) {
				adding_table = true;
				file1.WriteLine ("<table border=\"1\">");
			}
			// NewMessage.Show("Unimplemented");   
			string[] Cols = sText.Split (new string[1] { "||" }, StringSplitOptions.RemoveEmptyEntries);
			if (Cols != null && Cols.Length > 0) {
				file1.WriteLine ("<tr>");
				foreach (string s in Cols) {
					if (String.Empty != s && " " != s) {
						file1.WriteLine ("<td>");
						adding_row_to_table = true;
						FormatRestOfText (s);
						adding_row_to_table = false;
						
						file1.WriteLine ("</td>");
					}
				}
				file1.WriteLine ("</tr>");
			}
			
		}
		
		protected override void AddBullets (string sText)
		{
			AddAnyBullet (sText, "<ul>", "</ul>", "*");
			/*
            if ("" == startnormalbullet)
            {
                
                file1.Write("<ul>");
                startnormalbullet = "</ul>";
            }
            
            sText = sText.TrimStart('*').Trim();
            file1.Write(String.Format("<li>"));

            // have to do this to allow formating to appear on the line
            adding_bullet_now = true;
            FormatRestOfText(sText);
            adding_bullet_now = false;
            file1.Write("</li>");

            */
			
		}
		
		string startnormalbullet = "";
		bool adding_bullet_now = false;
		int highestbulletlevel = 0;

		public static int CountCharacters (string source, char match)
		{
			return CountCharacters (source, match, false);
		}
		/// <summary>
		/// Counts the characters.
		/// </summary>
		/// <returns>
		/// The characters.
		/// </returns>
		/// <param name='source'>
		/// Source.
		/// </param>
		/// <param name='match'>
		/// Match.
		/// </param>
		/// <param name='continuous'>
		/// If set to <c>true</c> continuous.
		/// </param>
		public static int CountCharacters (string source, char match, bool continuous)
		{
			int count = 0;
			bool countstarted = false;

			for (int i = 0; i < source.Length; i++) {
							

				char c = source [i];
				if (c == match) {
								

					countstarted = true;
					count++;
				} else {
					if (countstarted && continuous) {
						// we have hit a NON-match and are doing a continous search
						// that means our count is finished.
						break;
					}
				}
			}
			return count;
		}

//		/// <summary>
//		/// Got 2 levels of bullets working -- will have to generalize for more
//		/// </summary>
//		/// <param name="sText"></param>
//		/// <param name="sStartCode"></param>
//		/// <param name="sEndCode"></param>
//		/// <param name="sBulletSymbol"></param>
//		private void AddAnyBullet (string sText, string sStartCode, string sEndCode, string sBulletSymbol)
//		{
//			AddAnyBulletOld(sText, sStartCode, sEndCode, sBulletSymbol); return;
//			if ("" == startnormalbullet) {
//				
//				file1.Write (sStartCode);
//				startnormalbullet = sEndCode;
//			}
//			int CountOfBullets = -1;
//			if (sText.StartsWith (sBulletSymbol))
//			{
//				// we must have at least one * at the start to bother counting.
//				CountOfBullets = CountCharacters (sText, sBulletSymbol [0], true);
//				if (CountOfBullets > 3)
//					CountOfBullets = 3;
//			}
//			if (CountOfBullets > 0) {
//
//				file1.Write (String.Format ("Count = {0}, HighestbulletLevel = {1}", CountOfBullets, highestbulletlevel));
//				while (highestbulletlevel > CountOfBullets) {
//					// add /ol
//					file1.Write (sEndCode);
//					highestbulletlevel--;
//				}
//
//				// we only right a start tag if we are the first bullet on this level
//				if (CountOfBullets != highestbulletlevel && CountOfBullets > 1) {
//
//					file1.Write (sStartCode);
//				}
//				highestbulletlevel = CountOfBullets;
//			}
//
////			if (sText.StartsWith (sBulletSymbol + sBulletSymbol + sBulletSymbol)) {
////				if (3 != highestbulletlevel) {
////					file1.Write (sStartCode);
////				}
////				highestbulletlevel = 3;
////			} else
////				if (sText.StartsWith (sBulletSymbol + sBulletSymbol)) {
////
////				while (highestbulletlevel > 2) {
////					// add /ol
////					file1.Write (sEndCode);
////					highestbulletlevel--;
////				}
////
////
////				if (2 != highestbulletlevel) {
////					file1.Write (sStartCode);
////				}
////				highestbulletlevel = 2;
////			} else if (sText.StartsWith (sBulletSymbol)) {
////				// set back to 1
////				// because we were at a higher number and now we are a lower number
////				while (highestbulletlevel > 1) {
////					// add /ol
////					file1.Write (sEndCode);
////					highestbulletlevel--;
////				}
////				highestbulletlevel = 1;
////			}
//			sText = sText.TrimStart (sBulletSymbol [0]).Trim ();
//
//			// Hotfix May 2013
//			// By adding an empty * to the end of a list we can have a list end, without text between it and the ntext section
//			// previoulsy this would mix up the tags. 
//			if (sText != "") {
//				file1.Write (String.Format ("<li>"));
//			
//				// have to do this to allow formating to appear on the line
//				adding_bullet_now = true;
//				FormatRestOfText (sText);
//				adding_bullet_now = false;
//				file1.Write ("</li>");
//			}
//		}
		private void MinorHandlerForBullets (int Counted, string sStartCode, string sEndCode)
		{
			// set back to 1
			// because we were at a higher number and now we are a lower number
			while (highestbulletlevel >Counted) {
				// add /ol
				file1.Write (sEndCode);
				highestbulletlevel--;
			}
			if (Counted != highestbulletlevel && Counted != 1) {
				file1.Write (sStartCode);
			}
			highestbulletlevel = Counted;
		}

		// 08/05/2014 -- original, working to level 3 code before I tried going to infinite
		private void AddAnyBullet (string sText, string sStartCode, string sEndCode, string sBulletSymbol)
		{
			if ("" == startnormalbullet) {
				
				file1.Write (sStartCode);
				startnormalbullet = sEndCode;
			}
			int Counted = CountCharacters (sText, sBulletSymbol [0], true);
			//	if (sText.StartsWith (sBulletSymbol + sBulletSymbol + sBulletSymbol)) {

			if (Counted > 0) {
				MinorHandlerForBullets(Counted, sStartCode, sEndCode);
//				if (Counted == 4) {
//					MinorHandlerForBullets(Counted, sStartCode, sEndCode);
//				}
//				else
//
//				if (Counted == 3) {
//					MinorHandlerForBullets(Counted, sStartCode, sEndCode);
////					while (highestbulletlevel > Counted) {
////						// add /ol
////						file1.Write (sEndCode);
////						highestbulletlevel--;
////					}
////
////					if (Counted != highestbulletlevel) {
////						file1.Write (sStartCode);
////					}
////					highestbulletlevel = Counted;
//				} else if (Counted == 2) {
//					MinorHandlerForBullets(Counted, sStartCode, sEndCode);
//				
////					while (highestbulletlevel > Counted) {
////						// add /ol
////						file1.Write (sEndCode);
////						highestbulletlevel--;
////					}
////				
////				
////					if (2 != highestbulletlevel) {
////						file1.Write (sStartCode);
////					}
////					highestbulletlevel = Counted;
//				} else if (Counted == 1) {
//			
//					MinorHandlerForBullets(Counted, sStartCode, sEndCode);
////					// set back to 1
////					// because we were at a higher number and now we are a lower number
////					while (highestbulletlevel >Counted) {
////						// add /ol
////						file1.Write (sEndCode);
////						highestbulletlevel--;
////					}
////					highestbulletlevel = Counted;
//				}

			}

			sText = sText.TrimStart (sBulletSymbol [0]).Trim ();
			
			// Hotfix May 2013
			// By adding an empty * to the end of a list we can have a list end, without text between it and the ntext section
			// previoulsy this would mix up the tags. 
			if (sText != "") {
				file1.Write (String.Format ("<li>"));
				
				// have to do this to allow formating to appear on the line
				adding_bullet_now = true;
				FormatRestOfText (sText);
				adding_bullet_now = false;
				file1.Write ("</li>");
			}
		}
		/// <summary>
		/// Downt he road I probably want multileleve
		/// http://www.w3.org/TR/html401/struct/lists.html
		/// </summary>
		/// <param name="sText"></param>
		protected override void AddNumberedBullets (string sText)
		{
			
			AddAnyBullet (sText, "<ol>", "</ol>", "#");
			// NewMessage.Show("Unimplemented"); 
			
		}
		
		protected override void AddTableOfContents ()
		{
			// not implemented
		}
		
		public override void OverrideSceneBreak()
		{
			// The DOUBLE </p></p> is intentional in here. I need to wrap the element in a <p> and but </p> gets stripped. The doubl e</p> tricks my logic.
			//<p class=\"center\"> <p class=\"right\"></p>    
			FormatRestOfText(String.Format ("<p class=\"center\"><img {0} alt=\"section break\"/></p><p/><p>", this.controlFile.Overridesectionbreak), false);
		}
		/// <summary>
		/// Align text until next alignment command is hit
		/// 0 - Center
		/// 1 - Left
		/// 2 - Right
		/// </summary>
		/// <param name="nAlignment"></param>
		protected override void AlignText (int nAlignment)
		{
			// May 2013 - changed this to a counter because someties we have alignments happening really tight together
			NeedParagraphClosingCounter++;
			//NeedParagraphClosing = true;
			switch (nAlignment) {
			// February 2013 - does every P break require a closing because of the pattern I've already established?
			// May 2013 - these don't seem to be causing recent crop of errors [Nope: rejected. We cover this already with NeedParagraph Closing
			case 0:
				if (this.controlFile.ArealValidatorSafe_Align)
				{
					file1.WriteLine ("<p class=\"center\">");
				}
				else
				{
					file1.WriteLine ("<p align=\"center\">");
				}
				TabParagraph++;
				break;
			case 1:
				if (this.controlFile.ArealValidatorSafe_Align)
				{
					file1.WriteLine ("<p class=\"left\">");
				}
				else
				{
					file1.WriteLine ("<p align=\"left\">");
				}
				TabParagraph++;
				break;  
			case 2:
				if (this.controlFile.ArealValidatorSafe_Align)
				{
					file1.WriteLine ("<p class=\"right\">");
				}
				else
				{
					file1.WriteLine ("<p align=\"right\">");
				}
				TabParagraph++;
				break;
			}
		}
		
		protected override void InlineStrikeThrough (int nValue)
		{
			if (nValue > 0) {
				IsFormating = true;
				InlineWrite ("<strike>");
			} else {
				InlineWrite ("</strike>");
				IsFormating = false;
			}
		}

		private bool IsFormating = false;
		/// <summary>
		/// nvalue is ignored for underline
		/// </summary>
		/// <param name="nValue"></param>
		protected override void InlineUnderline (int nValue)
		{
			if (nValue > 0) {
				IsFormating = true;
				InlineWrite ("<u>");

			} else {
				InlineWrite ("</u>");
				IsFormating = false;
			}
		}
		
		/// <summary>
		/// nvalue is ignored for underline
		/// </summary>
		/// <param name="nValue"></param>
		protected override void InlineSuper (int nValue)
		{
			if (nValue > 0) {
				IsFormating = true;
				InlineWrite ("<sup>");
			} else {

				InlineWrite ("</sup>");
				IsFormating = false;
			}
		}
		
		/// <summary>
		/// nvalue is ignored for underline
		/// </summary>
		/// <param name="nValue"></param>
		protected override void InlineSub (int nValue)
		{
			if (nValue > 0) {
				IsFormating = true;
				InlineWrite ("<sub>");
			} else {
				InlineWrite ("</sub>");
				IsFormating = false;
			}
		}
		
		protected override void AddBookmark (string sBookmark)
		{
			sBookmark = sBookmark.Trim ();
			// decided not to show anchor text THOUGH we might want the {1} displayed for the footnote (but not the return)
			InlineWrite ("<a name=\"" + sBookmark + "\"></a>");
			
			
		}
		/// <summary>
		/// adds a picture
		/// </summary>
		/// <param name='nValue'>
		/// 
		/// </param>
		/// <param name='sPathToFile'>
		/// S path to file.
		/// </param>
		protected override void AddPicture (string sPathToFile)
		{
			//base.AddPicture (sPathToFile);

			string FileName = new FileInfo (sPathToFile).Name;
			if (FileName != Constants.BLANK) {
				string newDirectory = Path.Combine (sDirectory, "oebps");
				// first we copy the file to the location
				File.Copy (sPathToFile, Path.Combine (newDirectory, FileName));
				// just in root directory we just set the reference to the file directly
				InlineWrite ("<img src=\"" + FileName + "\"" + " alt=\"" + FileName + "\"/>");
			}
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sPathToFile"></param>
		protected override void AddLink (string sPathToFile, string sTitle)
		{
			string oTextToShow = sTitle;
			string oLink = sPathToFile;
			InlineWrite ("<a href=\"" + oLink + "\">" + oTextToShow + "</a>");
			
		}

		private List<string> listOfChapterNames = null;
		
		/// <summary>
		/// To do: Make AddTitle part of base
		/// </summary>
		/// <param name="sText"></param>
		protected override string AddTitle (string sText)
		{
			string storyidentifier = "";
			if (sText.IndexOf ("[[story]]") > -1) {
				storyidentifier = StoryIdentifier = "1_1*1";;
			}

			sText = base.AddTitle (sText);

			if (sText.IndexOf (":") > -1) {
				// creating fancier looking title if in format of 
				/*--- See Harry Potter book, can I format Chaptesr like that

 Chapter 1: Dog Eat Dog becomes

  -- Chapter 1 --
<h1>Dog Eat Dog</h1>
*/	
				string[] parts = sText.Split (new char[1] {':'}, StringSplitOptions.RemoveEmptyEntries);
				if (parts != null && parts.Length > 1)
				{
					if (null == listOfChapterNames) {
						listOfChapterNames = new List<string>();
					}
					// 14/07/2014 - we store chapter names with index [0] = Chapter 1, [1] = Chapter 2, et cetera
					// these will be used later when the TOC is generated to give Chapter names in the TOC.
					listOfChapterNames.Add (storyidentifier+parts[1]);

					InlineWrite (String.Format ("<p class=\"titlecenter\"> &#8212; {0} &#8212;</p>", parts[0]));
					InlineWrite (String.Format ("<h1>{0}</h1>", parts[1]));
				}
			}
			else
				InlineWrite (String.Format ("<h1>{0}</h1>", sText));

			return sText;
			
		}
		
		Hashtable FootnoteChapterIAmInHash;
		
		/// <summary>
		/// adding the link to the footnote whcih will be elsewhere
		/// </summary>
		/// <param name="sID"></param>
		protected override void AddFootnote (string sID)
		{
			sID = sID.Trim ();
			// we simply add the LINK
			string link = String.Format ("<a name=\"back{0}\"></a><a href=\"footnotes.xhtml#{0}\"><sup>(NOTE)</sup></a>", sID);
			
			// for it to work on Stanza I think the filenames need underliens and not spaces
			// off by one error with chapter (it is set to 'next' chaper) so we reset it and put it back
			chapter = chapter - 1;
			string sFile = chaptertoken.Replace (" ", "_").Trim () + ".xhtml";
			chapter = chapter + 1;
			// we add to a second lookup hash that jsut contains Key|Chapter that can be used when the footnotes are written up to create links to that chapter
			FootnoteChapterIAmInHash.Add (sID, sFile);
			
			InlineWrite (link);
			footnotesincurrentsection.Add (sID);
			
		}
		
		
		
		
		/// <summary>
		/// Writes out the footnote indicated
		/// </summary>
		/// <param name="sID"></param>
		protected  void AddActualFootnote (string sID)
		{
			
			
			sID = sID.Trim ();
			if (FootnoteHash.ContainsKey (sID) == true) {
				
				
				string sSource = (string)FootnoteHash [sID];
				
				sSource = sSource.Replace ("<br>", "<br></br>");
				// process text lienfeed
				
				object sText = sSource;
				object id = (object)sID;
				
				string chapterfile = "";
				if (FootnoteChapterIAmInHash.ContainsKey (sID) == true) {
					// now we grab a full chapter like chapter_7.xhtml
					chapterfile = FootnoteChapterIAmInHash [sID].ToString ();
					
				}
				
				string link = String.Format ("<span id=\"{0}\" class=\"footnote\">{1}<a class=\"EndNoteBackLink\" href=\"{2}#back{0}\">  <img class=\"Return\" alt=\"Return to Link Button\" src=\"images/return.png\" /></a></span><br></br>", sID, sText, chapterfile);
				
				
				
				if (controlFile.ShowFootNoteChapter == true) {
					link = " (" + chapterfile + ")  " + link;
				}
				
				InlineWrite (link);
			} else {
				NewMessage.Show (String.Format ("{0} footnote not found!", sID));
			}
			//   remember we need to have fancy formating and a return link and a Large name like NOTE that can be clicked on ipod
		}
	
		protected override void OnTitleChange ()
		{
			// here we break off and start a new file
			CloseCurrentFile ();
			currentFileBeingWritten = Path.Combine (sDirectory, String.Format ("oebps\\Chapter_{0}.xhtml", chapter.ToString ()));
			if (chapter == 9999) {
				currentFileBeingWritten = Path.Combine (sDirectory, String.Format ("oebps\\{0}.xhtml", "endnote"));
			}
			StartNewFile (currentFileBeingWritten);
		}

		bool WeHadABold = false;

		protected override void InlineBold (int nValue)
		{
			if (nValue > 0) {
				InlineWrite ("<b>");
				WeHadABold = true;
			} else {
				if (true == WeHadABold) {
					WeHadABold = false;
					//NewMessage.Show("Closing bold");
					InlineWrite ("</b>");
				}
			}
		}
		
		protected override void InlineItalic (int nValue)
		{
			if (nValue > 0) {
				IsFormating = true;
				InlineWrite ("<em>");
			} else {
				InlineWrite ("</em>");
				IsFormating = false;
			}
			
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sText"></param>
		/// <param name="nLevel">Heading level 1..4</param>
		protected override void AddHeader (string sText, int nLevel)
		{
			
			
			
			
			
			
			string slevel = "1";
			switch (nLevel) {
			case 1:
				slevel = "1";
				;
				break;
			case 2:
				slevel = "2";
				;
				break;
			case 3:
				slevel = "3";
				;
				break;
			case 4:
				slevel = "4";
				;
				break;
			case 5:
				slevel = "5";
				;
				break;
			}
			InlineWrite (String.Format ("<h{0}>{1}</h{0}>", slevel, sText));
			
		}
		
		protected override void AddPageBreak ()
		{
			InlineWrite ("<hr></hr>");
		}
		
		
		
		/// <summary>
		/// at end?
		/// </summary>
		protected override void Cleanup ()
		{
//			// close off any paragraph tabs
			// May 2013 - this breaks more than it fixes
//			for (int i = 1; i <= TabParagraph; i++)
//			{
//				//March 2013 - trying to remove  
//				// May 2013 - I added this BACK in to try to counter ALIGN tags. The only PLACE this is increment is inside of Alignments.
//				InlineWrite("</p>");
//			}
//			TabParagraph = 0;

			// May 2013 - removing this because we call CloseCurrentFile oruselves -- which is also called in base.Cleanup.
			//	base.Cleanup ();
			// moved FOOTNOTE CODE to cleanup to write out at end of file


			
			try {
				if (footnotesincurrentsection.Count > 0) {
					// write out footnotes.xhtml
					string footnotes = Path.Combine (sDirectory, String.Format ("oebps\\footnotes.xhtml", chapter.ToString ()));
					StartNewFile (footnotes);


					InlineWrite ("<hr></hr><h3>Footnotes</h3>");
				
					// write out footnotes for current section
					int count = 0;
					foreach (string s in footnotesincurrentsection) {
						count++;
						InlineWrite ("<strong>[" + count.ToString () + "] </strong>");
						AddActualFootnote (s);
					}
				}
			
				CloseCurrentFile ();
			} catch (Exception ex) {
				NewMessage.Show (ex.ToString ());
			}
			CreateContentOPF ();




			if (File.Exists (ZIP7) == true) {
				NewMessage.Show ("01/04/2014 -- The ZIP worked but if you overwrote the style sheet, the changes did not appear unless you manually Compressed the directory yourself and added the .epub extension. So since I had to do it by hand I removed this.");
			}
			//01/04/2014
//			if (File.Exists (ZIP7) == true) {
//				// now we write out a batch file
//				string zipfile = Path.Combine (directory_to_sourcefiles, "lastzip.bat");
//				string sourcepath = sDirectory;
//				string zippath = sDirectory + ".epub";
//				///<img alt="--" title="" src="images/fleuron.png" width="275" height="20"></img>
//				StreamWriter zip = new StreamWriter (zipfile);
//				string scommand = String.Format ("\"{0}\" a -tzip \"{1}\" \"{2}\"", ZIP7, zippath, sourcepath);
//				zip.WriteLine (scommand);
//				zip.Close ();
//				// runbatch
//				General.OpenDocument (zipfile, "");
//			}
			
			if (false == SuppressMessages) {
				NewMessage.Show ("Done!");
			}
			/* // update the table of contents if present
            if (oDoc.TablesOfContents.Count > 0)
            {
                oDoc.TablesOfContents[1].Update();
            }

            if (controlFile.Linespace != -1)
            {
                // do cleanup
                // fix line spacing
                foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in oDoc.Paragraphs)
                {
                    paragraph.Format.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;
                    paragraph.Format.LineSpacing = controlFile.Linespace;
                }
            }

            */
			
		}
		/// <summary>
		/// Raises the finished totally event.
		/// 
		/// This canbe used for post-processing, like iterating through "files"
		/// </summary>
		protected override void OnFinishedTotally ()
		{
			base.OnFinishedTotally ();
		

		

			//return;
			if (this.controlFile.EpubRemoveDoublePageTags == true || this.controlFile.NovelMode == true) {
				foreach (string aFileUsed in this.files) {
					// process the file

					string SourceFile = Path.Combine (controlFile.OutputDirectory + "\\" + GetDateDirectory, "oebps\\" + aFileUsed);

					//NewMessage.Show (SourceFile);
					//rename file
					if (File.Exists (SourceFile)) {
						System.IO.File.Move (SourceFile, SourceFile + "2");

						// open filesource
						StreamReader source = new StreamReader (SourceFile + "2");
					

						// openfiledest with name of old file
						StreamWriter writer = new StreamWriter (SourceFile);
						// copy
						string sLine = source.ReadLine ();
						while (sLine != null) {


						


							if (this.controlFile.EpubRemoveDoublePageTags == true) {
								sLine = sLine.Replace ("<p></p>", "<br/>");
								// if there's a bunch of breaks together consolidate them
								sLine = sLine.Replace ("<br/><br/><br/><br/><br/>", "<br/>");
								sLine = sLine.Replace ("<br/><br/><br/>", "<br/>");
							} else
							if (this.controlFile.NovelMode == true) {
								//19/06/2014 - blank open line shoudl not cause error
								//sLine = sLine.Replace ("<body>\r" , "AT");
								if (sLine == "<body>")
								{
									writer.WriteLine (sLine);
									// we immediately read the next line and remove a <p></p> at start. Yes. I know. Dangerous.
									sLine = source.ReadLine ();
									// to ensure acfuracy, only grab start of line
									try
									{

										string sub = sLine.Substring(0, 7);
										if (sub == "<p></p>")
										{
											sLine = sLine.Remove(0, 7);
										}
										sub = sLine.Substring (0,4);
										if (sub == "</p>")
										{
											sLine = sLine.Remove(0, 4);
										}
									}
									catch (Exception)
									{
									}
									
								}
							
								//19/06/2014 - Added this because of a problem with POOL
								sLine = sLine.Replace ("</p>.<p></p>", ".</p>");

								//19/06/2014 -- if we have a [[fact], in middle of a line RIGHT AFTER A SCENE BREAK, the file breaks.
								sLine = sLine.Replace ("</p>,", ",");
								sLine = sLine.Replace ("</p>.", ".");
								sLine = sLine.Replace ("</p>!", "!");
								sLine = sLine.Replace ("</p>?", "?");
								sLine = sLine.Replace ("</p> -", " -");
								sLine = sLine.Replace ("</p> ", " "); // this might be pushing my luck, but it does break (see Novel-factafterscene_WITHOUTCOMMAN
								sLine = sLine.Replace (" </p>", " "); // this might be pushing my luck, but it does break (see Novel-factafterscene_WITHOUTCOMMAN


								//19/06/2014 - opening a titled section causes errors
								sLine = sLine.Replace ("</h1><p></p><div", "</h1><div");
//								<body>
//									<p></p>	


								// older ones
								sLine = sLine.Replace ("\t", "<p>");
								sLine = sLine.Replace ("<p></p>", "</p>");



								sLine = sLine.Replace ("</p></p></body>", "</p></body>");
								sLine = sLine.Replace ("<div class=\"quote\"></p>", "</p><div class=\"quote\"><p>");

								sLine = sLine.Replace ("</div></p><p>", "</div><p>");
								sLine = sLine.Replace ("#</p></p><p align=", "#</p><p align=");
								sLine = sLine.Replace ("#</p></p><p class=", "#</p><p class=");
								sLine = sLine.Replace ("</p></p><p>", "</p><p>");
								//  33 - <p><p>Watts</p></div></body>
								sLine = sLine.Replace ("<p><p>", "<p>");

								//need to test these fixes C27
								sLine = sLine.Replace ("</p><em>", "<em>");

								// test c8
								sLine = sLine.Replace ("<p></p><p>", "</p><p>");
								//test c4 / 20
								sLine = sLine.Replace ("</p><hr></hr></p>", "<hr></hr>");
								//c 30 
								sLine = sLine.Replace ("<p></p></body>", "</p></body>");
								//c21
								sLine = sLine.Replace ("</p></p></body>", "</p></body>");
								//sLine = sLine.Replace("</p>","</p>");
								//c35
								sLine = sLine.Replace ("<div class=\"past\"></p>", "</p><div class=\"past\">");


								// keep very end, apply other rules first -- Preface
								while (sLine.IndexOf ("</p></p></p></p></p></p></p></p></p></p></p>") > -1)
								{
									sLine = sLine.Replace ("</p></p></p></p></p></p></p></p></p></p></p>", "");
								}

								// keep very end, apply other rules first -- Preface
								while (sLine.IndexOf ("</p></p></p></p></p></p></p></p></p></p>") > -1)
								{
									sLine = sLine.Replace ("</p></p></p></p></p></p></p></p></p></p>", "");
								}

								// keep very end, apply other rules first -- Preface
								while (sLine.IndexOf ("</p></p></p></p></p></p></p></p></p>") > -1)
								{
									sLine = sLine.Replace ("</p></p></p></p></p></p></p></p></p>", "");
								}

								while (sLine.IndexOf ("</p></p></p></p></p>") > -1)
								{
									sLine = sLine.Replace ("</p></p></p></p></p>", "");
								}
									sLine = sLine.Replace ("</p></p></p>", "");
									sLine = sLine.Replace ("</p></p>", "");


								if (this.controlFile.CustomEllip) {
									//NewMessage.Show ("h");

									//decided an end ellips would have a space before but not after
									sLine = sLine.Replace("....", "&hellip;.");
									// mid-line ellipsis is suppose to have spce before and after
									sLine = sLine.Replace("...", "&hellip; ");
								}
							}
							writer.WriteLine (sLine);
							sLine = source.ReadLine ();

						
						}
						//writer.WriteLine("hi");
						//close
						source.Close ();
						writer.Close ();
						// delete source
						File.Delete (SourceFile + "2");
					}
				}
			}

		}
		string StoryIdentifier = "1_1*1";
		/// <summary>
		/// goes through the files aaraylist and creates an xml file content.opf
		/// </summary>
		private void CreateContentOPF ()
		{
			
			if (files != null && files.Count > 0) {
				
				/*// we remove footer.xhtml (june 2010 to put it at end)
				string footnote = (files[0].ToString());
				files.Remove(files[0]);
				files.Add(footnote);
				*/
					
				// write content.opf, reading initial stuff.
				StreamReader reader = new StreamReader (Path.Combine (directory_to_sourcefiles, "oebps\\template_content.opf"));
				StreamWriter writer = new StreamWriter (Path.Combine (sDirectory, "oebps\\content.opf"));
				
				
				string line = reader.ReadLine ();
				while (line != null) {
					
					
					line = ParseLineForId (line);
					writer.WriteLine (line);
					
					line = reader.ReadLine ();
					if (line == "[chaptersstart]") {
						// now start inserting chapters
						foreach (string s in files) {
							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s.Replace (" ", "_").Trim ();
							writer.WriteLine (String.Format ("<item id=\"{0}\" href=\"{0}\" media-type=\"application/xhtml+xml\"/>", sFile));
						}
						// we read the next line
						line = reader.ReadLine ();
					}
					if (line == "[tocstart]") {

						foreach (string s in files) {


							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s.Replace (" ", "_").Trim ();

							// this is basically dicating the order that the book appears
							writer.WriteLine (String.Format ("<itemref idref=\"{0}\"/>", sFile));
						}
						line = reader.ReadLine ();
					}
					
				}
				reader.Close ();
				writer.Close ();

				//
				// now we write the toc blank.ncx file
				//
				reader = new StreamReader (Path.Combine (directory_to_sourcefiles, "oebps\\toc blank.ncx"));
				writer = new StreamWriter (Path.Combine (sDirectory, "oebps\\toc.ncx"));
				line = null;
				
				line = reader.ReadLine ();
				while (line != null) {
					// TO DO parse IDs from the ID database
					line = ParseLineForId (line);
					writer.WriteLine (line);
					
					line = reader.ReadLine ();
					
					if (line == "[chaptersstart]") {
						int count = 2;
						foreach (string s in files) {  //make navlabel nice and translate the ids
							string navlabel = s.Replace (".xhtml", "").Trim ();
							//26/04/2014 - don't want underscores in display name
							navlabel = navlabel.Replace ("_", " ").Trim ();
							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s.Replace (" ", "_").Trim ();
							
							writer.WriteLine (String.Format ("<navPoint id=\"{0}\" playOrder=\"{1}\"><navLabel><text>{2}</text></navLabel><content src=\"{0}\"/></navPoint>",
							                               sFile, count, navlabel));
							count++;
						}
						line = reader.ReadLine ();
					}
				}
				reader.Close ();
				writer.Close ();
				
			

			
				//
				// now we write the actual TOC file that is inserted at front of book
				//
				reader = new StreamReader (Path.Combine (directory_to_sourcefiles, "oebps\\template_toc.xhtml"));
				writer = new StreamWriter (Path.Combine (sDirectory, "oebps\\toc.xhtml"));
				line = null;
			
				line = reader.ReadLine ();
				while (line != null) {
					// TO DO parse IDs from the ID database
					line = ParseLineForId (line);
					writer.WriteLine (line);
				
					line = reader.ReadLine ();
				
					if (line == "[chaptersstart]") {
						int count = 2;
						int ChapterCountIndex = controlFile.ChapterTitleOffset;
						foreach (string s in files) {  //make navlabel nice and translate the ids


						

							string navlabel = s.Replace (".xhtml", "").Trim ();
							navlabel = navlabel.Replace ("_", " ").ToUpper ();
							ChapterCountIndex++; // this index is used to figure out which chapter name (if any) to add
							try
							{
								if (listOfChapterNames != null && ChapterCountIndex < listOfChapterNames.Count && ChapterCountIndex >= 0)
								{

									// tweaking things so that short stories don't show the word "Chapter" in TOC
									string chapterName = listOfChapterNames[ChapterCountIndex];
									if (chapterName.IndexOf(StoryIdentifier) > -1)
									{
										chapterName = chapterName.Replace (StoryIdentifier, "");
										navlabel = String.Format ("{0}", chapterName);
									}
									else
										navlabel = String.Format ("{0} - {1}", navlabel, chapterName);
								}
							}
							catch (System.Exception ex)
							{
								NewMessage.Show (ex.ToString());
							}
							
							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s;// s.Replace(" ", "_").Trim();
							writer.WriteLine (String.Format ("<p><a href=\"{0}\">{1}</a></p>",
						                               sFile, navlabel));


							count++;
						}
						line = reader.ReadLine ();
					}
				}
				reader.Close ();
				writer.Close ();
		
			
			}
			
		}
		/// <summary>
		/// goes through the current line being read and replaces tags like [title] with ids, if they exist.
		/// This happens to template_content.opf and toc blank.ncx
		/// </summary>
		/// <param name="line"></param>
		/// <returns></returns>
		private string ParseLineForId (string line)
		{
			// this outer if is just to save some performance time
			if (line.IndexOf ("[") > -1) {
				foreach (string key in ids.Keys) {
					string label = String.Format ("[{0}]", key);
					line = line.Replace (label, ids [key].ToString ());
				}
			}
			return line;
		}
		
		/// <summary>
		/// For children like sendePub will store in a hash for use when generating ePub files
		/// </summary>
		/// <param name="id"></param>
		/// <param name="idtext"></param>
		protected override void AddId (string id, string idtext)
		{
			if (ids.ContainsKey (id) == false) {
				ids.Add (id, idtext);
			}
		}

		public override string ToString ()
		{
			return string.Format ("[sendePub]");
		}
	}
}
