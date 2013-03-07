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
		string directory_to_sourcefiles ="";// @"C:\Users\Brent\Documents\Visual Studio 2005\Projects\SendTextAway\SendTextAway\bin\Debug\epubfiles\";
		
		private bool NeedParagraphClosing; // used for centering. If true the next InlineWrite will write a </p> closing tag
		private string sDirectory = "";
		
		ArrayList footnotesincurrentsection; // footnotes can be chapter, or if I can figure out a good solution the entire document
		
		private Hashtable ids;
		
		
		public static string GetDateDirectory {
			get { return DateTime.Today.ToString("yyyy MM dd") ;}
		}
		
		/// <summary>
		/// before whatever initial operations are required to open the filestream
		/// or whatever (in the case of Word Auto, will require global variables)
		/// </summary>
		protected override int InitializeDocument (ControlFile _controlFile)
		{
			directory_to_sourcefiles = _controlFile.TemplateDirectory;
			ZIP7 = _controlFile.Zipper;
			
			if (null == ZIP7 || Constants.BLANK == ZIP7 || null == directory_to_sourcefiles || Constants.BLANK == directory_to_sourcefiles) {
				NewMessage.Show (Loc.Instance.GetString ("To generate an epub you need to specificy a valid path to 7Zip and to the template files used to generate the final epub files"));
				return -1;
			}
			// create an output directory based on date
			sDirectory = Path.Combine(_controlFile.OutputDirectory, GetDateDirectory);
			Directory.CreateDirectory(sDirectory);
			File.Copy(Path.Combine(directory_to_sourcefiles, "mimetype"), Path.Combine(sDirectory, "mimetype"), true);
			
			Directory.CreateDirectory(Path.Combine(sDirectory, "oebps"));
			DirectoryInfo images = Directory.CreateDirectory(Path.Combine(sDirectory, "oebps\\images"));
			Directory.CreateDirectory(Path.Combine(sDirectory, "META-INF"));
			
			// start with footer as default first file?
			currentFileBeingWritten = Path.Combine(sDirectory,"oebps\\" + "preface.xhtml");
			
			chapter = 1;
			base.InitializeDocument(_controlFile);
			
			
			// copy required files
			
			// copy image files
			FileUtils.Copy(new DirectoryInfo(Path.Combine(directory_to_sourcefiles,"oebps\\images")), images, "*.*", "*.*", false, 2, new System.Windows.Forms.ProgressBar());
			
			
			File.Copy(Path.Combine(directory_to_sourcefiles,"oebps\\copyright.xhtml"), Path.Combine(sDirectory, "oebps\\copyright.xhtml"),true);
			File.Copy(Path.Combine(directory_to_sourcefiles, "oebps\\stylesheet.css"), Path.Combine(sDirectory, "oebps\\stylesheet.css"), true);
			File.Copy(Path.Combine(directory_to_sourcefiles, "oebps\\legal.xhtml"), Path.Combine(sDirectory, "oebps\\legal.xhtml"), true);
			File.Copy(Path.Combine(directory_to_sourcefiles, "oebps\\title_page.xhtml"), Path.Combine(sDirectory, "oebps\\title_page.xhtml"), true);
			File.Copy(Path.Combine(directory_to_sourcefiles, "oebps\\page-template.xpgt"), Path.Combine(sDirectory, "oebps\\page-template.xpgt"), true);
			
			
			File.Copy(Path.Combine(directory_to_sourcefiles,"META-INF\\container.xml"), Path.Combine(sDirectory, "META-INF\\container.xml"),true);
			footnotesincurrentsection = new ArrayList();
			FootnoteChapterIAmInHash = new Hashtable();
			ids = new Hashtable();
			
			
			return 1;
		}
		
		
		
		/// <summary>
		/// generic Insert function for InsertHeader and InsertFooter
		/// </summary>
		/// <param name="sTemplate"></param>
		/// <param name="sChapterToken"></param>
		private void InsertFromTemplate(string sTemplate, string sChapterToken)
		{
			if (null == file1)
			{
				throw new Exception("no file open for writing");
			}
			// loads a header file
			StreamReader header = new StreamReader(sTemplate);
			if (null != header)
			{
				string line = header.ReadLine();
				while (line != null)
				{
					if ("" != sChapterToken)
					{
						line = line.Replace("[chapter]", sChapterToken);
						
					}
					file1.WriteLine(line);
					line = header.ReadLine();
				}
			}
			header.Close();
			header.Dispose();
		}
		/// <summary>
		/// inserts text for header file
		/// <param name="sChapterToken">If not "" then if we find [chapter] we replace it with sChapterToken</param>
		/// </summary>
		protected override void InsertHeaderForNewFile(string sChapterToken)
		{
			InsertFromTemplate(Path.Combine(directory_to_sourcefiles, "header.txt"), sChapterToken);
			
		}
		
		protected override void InsertFooterorNewFile(string sChapterToken)
		{
			InsertFromTemplate(Path.Combine(directory_to_sourcefiles, "footer.txt"), sChapterToken);
		}
		
		/// <summary>
		/// current chapter file being closed off
		/// </summary>
		protected override void CloseCurrentFile()
		{
			
			// moved FOOTNOTE CODE to cleanup to write out at end of file
			
			
			
			// blank list of footnotes for this current section (MIGHT BE MOVED IF CHANGE FOOTNOTE SYSTEM)
			// June 2010 now writign to end file instead footnotesincurrentsection = new ArrayList(); ;
			
			
			// close off any paragraph tabs
			for (int i = 1; i <= TabParagraph; i++)
			{
				//March 2013 - trying to remove  
				//InlineWrite("</p>");
			}
			TabParagraph = 0;
			base.CloseCurrentFile();
			
			
			
			
			
		}
		private int TabParagraph = 0; // count how many <p> tagsa we've introduce through plain, unformatted text
		/// <summary>
		/// While processing linline formatting (bold, et cetera), this is used to write out a line of text
		/// </summary>
		/// <param name="sText"></param>
		protected override void InlineWrite(string sText)
		{
			if (true == adding_table && false == adding_row_to_table)
			{
				adding_table = false;
				file1.WriteLine("</table>");
			}
			
			
			// if we were writing a bullet list, close it
			if ("" != startnormalbullet && false == adding_bullet_now)
			{
				file1.Write(startnormalbullet);
				startnormalbullet = "";
			}



			sText = sText.Replace (Environment.NewLine, "<p></p>");


			string sLine = "{0}";



			// we only do a <p> break if we detect a tab
			if (sText.IndexOf("\t") > -1)
			{

				
				//NeedParagraphClosing = true;
				
			}
			
			
			if (true == NeedParagraphClosing && "" != sText )
			{
				sLine = "{0}</p>";
				NeedParagraphClosing = false;
			}
			
			sLine = String.Format(sLine, sText);
			
			file1.Write(sLine);
			
			
			//oSelection.TypeText(sText);
		}
		
		private string lastmultiline = "";
		protected override bool AddStartMultiLineFormat(string sFormatType)
		{
			InlineWrite("<div class=\"" + sFormatType + "\">");
			lastmultiline = sFormatType;
			return true;
		}
		
		/// <summary>
		/// resets the formatting of a multi line format to normal format
		/// </summary>
		/// <returns></returns>
		protected override void AddEndMultiLineFormat()
		{
			if ("" != lastmultiline)
			{
				//InlineWrite("</"+lastmultiline+">");
				InlineWrite("</div>");
			}
		}
		
		bool adding_table;
		bool adding_row_to_table; // while adding row we don't put a /table
		protected override void AddTable(string sText)
		{
			if (false == adding_table)
			{
				adding_table = true;
				file1.WriteLine("<table border=\"1\">");
			}
			// NewMessage.Show("Unimplemented");   
			string[] Cols = sText.Split(new string[1] { "||" }, StringSplitOptions.RemoveEmptyEntries);
			if (Cols != null && Cols.Length > 0)
			{
				file1.WriteLine("<tr>");
				foreach (string s in Cols)
				{
					if (String.Empty != s && " " != s)
					{
						file1.WriteLine("<td>");
						adding_row_to_table = true;
						FormatRestOfText(s);
						adding_row_to_table = false;
						
						file1.WriteLine("</td>");
					}
				}
				file1.WriteLine("</tr>");
			}
			
		}
		
		
		protected override void AddBullets(string sText)
		{
			AddAnyBullet(sText, "<ul>", "</ul>", "*");
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
		
		/// <summary>
		/// Got 2 levels of bullets working -- will have to generalize for more
		/// </summary>
		/// <param name="sText"></param>
		/// <param name="sStartCode"></param>
		/// <param name="sEndCode"></param>
		/// <param name="sBulletSymbol"></param>
		private void AddAnyBullet(string sText, string sStartCode, string sEndCode, string sBulletSymbol)
		{
			if ("" == startnormalbullet)
			{
				
				file1.Write(sStartCode);
				startnormalbullet = sEndCode;
			}
			
			if (sText.StartsWith(sBulletSymbol + sBulletSymbol + sBulletSymbol))
			{
				if (3 != highestbulletlevel)
				{
					file1.Write(sStartCode);
				}
				highestbulletlevel = 3;
			}
			else
				if (sText.StartsWith(sBulletSymbol + sBulletSymbol))
			{
				if (2 != highestbulletlevel)
				{
				file1.Write(sStartCode);
				}
				highestbulletlevel = 2;
			}
			else if (sText.StartsWith(sBulletSymbol))
			{
				// set back to 1
				// because we were at a higher number and now we are a lower number
				while (highestbulletlevel > 1)
				{
					// add /ol
					file1.Write(sEndCode);
					highestbulletlevel--;
				}
				highestbulletlevel = 1;
			}
			sText = sText.TrimStart(sBulletSymbol[0]).Trim();
			file1.Write(String.Format("<li>"));
			
			// have to do this to allow formating to appear on the line
			adding_bullet_now = true;
			FormatRestOfText(sText);
			adding_bullet_now = false;
			file1.Write("</li>");
		}
		
		/// <summary>
		/// Downt he road I probably want multileleve
		/// http://www.w3.org/TR/html401/struct/lists.html
		/// </summary>
		/// <param name="sText"></param>
		protected override void AddNumberedBullets(string sText)
		{
			
			AddAnyBullet(sText, "<ol>", "</ol>","#");
			// NewMessage.Show("Unimplemented"); 
			
		}
		
		protected override void AddTableOfContents()
		{
			// not implemented
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
			NeedParagraphClosing = true;
			switch (nAlignment)
			{
				// February 2013 - does every P break require a closing because of the pattern I've already established?
			case 0: file1.WriteLine("<p align=\"center\">"); break;
			case 1: file1.WriteLine("<p align=\"left\">"); break;  
			case 2: file1.WriteLine("<p align=\"right\">"); break;
			}
		}
		
		protected override void InlineStrikeThrough(int nValue)
		{
			if (nValue > 0)
			{
				InlineWrite("<strike>");
			}
			else
			{
				InlineWrite("</strike>");
			}
		}
		
		/// <summary>
		/// nvalue is ignored for underline
		/// </summary>
		/// <param name="nValue"></param>
		protected override void InlineUnderline(int nValue)
		{
			if (nValue > 0)
			{
				InlineWrite("<u>");
			}
			else
			{
				InlineWrite("</u>");
			}
		}
		
		/// <summary>
		/// nvalue is ignored for underline
		/// </summary>
		/// <param name="nValue"></param>
		protected override void InlineSuper(int nValue)
		{
			if (nValue > 0)
			{
				InlineWrite("<sup>");
			}
			else
			{
				InlineWrite("</sup>");
			}
		}
		
		/// <summary>
		/// nvalue is ignored for underline
		/// </summary>
		/// <param name="nValue"></param>
		protected override void InlineSub(int nValue)
		{
			if (nValue > 0)
			{
				InlineWrite("<sub>");
			}
			else
			{
				InlineWrite("</sub>");
			}
		}
		
		protected override void AddBookmark(string sBookmark)
		{
			sBookmark = sBookmark.Trim();
			// decided not to show anchor text THOUGH we might want the {1} displayed for the footnote (but not the return)
			InlineWrite("<a name=\""+sBookmark+"\"></a>");
			
			
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sPathToFile"></param>
		protected override void AddLink(string sPathToFile, string sTitle)
		{
			string oTextToShow = sTitle;
			string oLink = sPathToFile;
			InlineWrite("<a href=\"" + oLink + "\">" + oTextToShow + "</a>");
			
		}
		
		/// <summary>
		/// To do: Make AddTitle part of base
		/// </summary>
		/// <param name="sText"></param>
		protected override string AddTitle(string sText)
		{
			sText = base.AddTitle(sText);
			
			InlineWrite(String.Format("<h1>{0}</h1>", sText));
			return sText;
			
		}
		
		Hashtable FootnoteChapterIAmInHash;
		
		/// <summary>
		/// adding the link to the footnote whcih will be elsewhere
		/// </summary>
		/// <param name="sID"></param>
		protected override void AddFootnote(string sID)
		{
			sID = sID.Trim();
			// we simply add the LINK
			string link = String.Format("<a name=\"back{0}\"></a><a href=\"footnotes.xhtml#{0}\"><sup>(NOTE)</sup></a>",sID);
			
			// for it to work on Stanza I think the filenames need underliens and not spaces
			// off by one error with chapter (it is set to 'next' chaper) so we reset it and put it back
			chapter = chapter - 1;
			string sFile = chaptertoken.Replace(" ", "_").Trim()+".xhtml";
			chapter = chapter + 1;
			// we add to a second lookup hash that jsut contains Key|Chapter that can be used when the footnotes are written up to create links to that chapter
			FootnoteChapterIAmInHash.Add(sID, sFile);
			
			InlineWrite(link);
			footnotesincurrentsection.Add(sID);
			
		}
		
		
		
		
		/// <summary>
		/// Writes out the footnote indicated
		/// </summary>
		/// <param name="sID"></param>
		protected  void AddActualFootnote(string sID)
		{
			
			
			sID = sID.Trim();
			if (FootnoteHash.ContainsKey(sID) == true)
			{
				
				
				string sSource = (string)FootnoteHash[sID];
				
				sSource = sSource.Replace("<br>", "<br></br>");
				// process text lienfeed
				
				object sText = sSource;
				object id = (object)sID;
				
				string chapterfile = "";
				if (FootnoteChapterIAmInHash.ContainsKey(sID) == true)
				{
					// now we grab a full chapter like chapter_7.xhtml
					chapterfile = FootnoteChapterIAmInHash[sID].ToString();
					
				}
				
				string link = String.Format("<span id=\"{0}\" class=\"footnote\">{1}<a class=\"EndNoteBackLink\" href=\"{2}#back{0}\">  <img class=\"Return\" alt=\"Return to Link Button\" src=\"images/return.png\" /></a></span><br></br>", sID, sText, chapterfile);
				
				
				
				if (controlFile.ShowFootNoteChapter == true)
				{
					link = " (" + chapterfile + ")  " + link ;
				}
				
				InlineWrite(link);
			}
			
			else
			{
				NewMessage.Show(String.Format("{0} footnote not found!", sID));
			}
			//   remember we need to have fancy formating and a return link and a Large name like NOTE that can be clicked on ipod
		}
		
		protected override void OnTitleChange()
		{
			// here we break off and start a new file
			CloseCurrentFile();
			currentFileBeingWritten = Path.Combine(sDirectory, String.Format("oebps\\Chapter_{0}.xhtml", chapter.ToString()));
			StartNewFile(currentFileBeingWritten);
		}
		bool WeHadABold = false;
		protected override void InlineBold(int nValue)
		{
			if (nValue > 0)
			{
				InlineWrite("<b>");
				WeHadABold = true;
			}
			else
			{
				if (true == WeHadABold)
				{
					WeHadABold = false;
					//NewMessage.Show("Closing bold");
					InlineWrite("</b>");
				}
			}
		}
		
		protected override void InlineItalic(int nValue)
		{
			if (nValue > 0)
			{
				InlineWrite("<em>");
			}
			else
			{
				InlineWrite("</em>");
			}
			
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sText"></param>
		/// <param name="nLevel">Heading level 1..4</param>
		protected override void AddHeader(string sText, int nLevel)
		{
			
			
			
			
			
			
			string slevel = "1";
			switch (nLevel)
			{
			case 1: slevel = "1"; ; break;
			case 2: slevel = "2"; ; break;
			case 3: slevel = "3"; ; break;
			case 4: slevel = "4"; ; break;
			case 5: slevel = "5"; ; break;
			}
			InlineWrite(String.Format("<h{0}>{1}</h{0}>", slevel, sText));
			
		}
		
		
		
		protected override void AddPageBreak()
		{
			InlineWrite("<hr></hr>");
		}
		
		
		
		/// <summary>
		/// at end?
		/// </summary>
		protected override void Cleanup ()
		{
			base.Cleanup ();
			
			
			// write out footnotes.xhtml
			string footnotes = Path.Combine (sDirectory, String.Format ("oebps\\footnotes.xhtml", chapter.ToString ()));
			StartNewFile (footnotes);
			
			if (footnotesincurrentsection.Count > 0) {
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
			
			CreateContentOPF ();
			
			if (File.Exists (ZIP7) == true) {
				// now we write out a batch file
				string zipfile = Path.Combine (directory_to_sourcefiles, "lastzip.bat");
				string sourcepath = sDirectory;
				string zippath = sDirectory + ".epub";
				///<img alt="--" title="" src="images/fleuron.png" width="275" height="20"></img>
				StreamWriter zip = new StreamWriter (zipfile);
				string scommand = String.Format ("\"{0}\" a -tzip \"{1}\" \"{2}\"", ZIP7, zippath, sourcepath);
				zip.WriteLine (scommand);
				zip.Close ();
				// runbatch
				General.OpenDocument (zipfile, "");
			}
			
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
		/// goes through the files aaraylist and creates an xml file content.opf
		/// </summary>
		private void CreateContentOPF()
		{
			
			if (files != null && files.Count > 0)
			{
				
				/*// we remove footer.xhtml (june 2010 to put it at end)
				string footnote = (files[0].ToString());
				files.Remove(files[0]);
				files.Add(footnote);
				*/
					
					// write content.opf, reading initial stuff.
					StreamReader reader = new StreamReader(Path.Combine(directory_to_sourcefiles, "oebps\\template_content.opf"));
				StreamWriter writer = new StreamWriter(Path.Combine(sDirectory, "oebps\\content.opf"));
				
				
				string line = reader.ReadLine();
				while (line != null)
				{
					
					
					line = ParseLineForId(line);
					writer.WriteLine(line);
					
					line = reader.ReadLine();
					if (line == "[chaptersstart]")
					{
						// now start inserting chapters
						foreach (string s in files)
						{
							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s.Replace(" ", "_").Trim();
							writer.WriteLine(String.Format("<item id=\"{0}\" href=\"{0}\" media-type=\"application/xhtml+xml\"/>", sFile));
						}
						// we read the next line
						line = reader.ReadLine();
					}
					if (line == "[tocstart]")
					{
						foreach (string s in files)
						{
							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s.Replace(" ", "_").Trim();
							// this is basically dicating the order that the book appears
							writer.WriteLine(String.Format("<itemref idref=\"{0}\"/>", sFile));
						}
						line = reader.ReadLine();
					}
					
				}
				reader.Close();
				writer.Close();
				
				// now we write the toc blank.ncx file
				reader = new StreamReader(Path.Combine(directory_to_sourcefiles, "oebps\\toc blank.ncx"));
				writer = new StreamWriter(Path.Combine(sDirectory, "oebps\\toc.ncx"));
				line = null;
				
				line = reader.ReadLine();
				while (line != null)
				{
					// TO DO parse IDs from the ID database
					line = ParseLineForId(line);
					writer.WriteLine(line);
					
					line = reader.ReadLine();
					
					if (line == "[chaptersstart]")
					{
						int count = 2;
						foreach (string s in files)
						{  //make navlabel nice and translate the ids
							string navlabel = s.Replace(".xhtml", "").Trim();
							// for it to work on Stanza I think the filenames need underliens and not spaces
							string sFile = s.Replace(" ", "_").Trim();
							
							writer.WriteLine(String.Format("<navPoint id=\"{0}\" playOrder=\"{1}\"><navLabel><text>{2}</text></navLabel><content src=\"{0}\"/></navPoint>",
							                               sFile, count, navlabel));
							count++;
						}
						line = reader.ReadLine();
					}
				}
				reader.Close();
				writer.Close();
				
			}
			
			
		}
		/// <summary>
		/// goes through the current line being read and replaces tags like [title] with ids, if they exist.
		/// This happens to template_content.opf and toc blank.ncx
		/// </summary>
		/// <param name="line"></param>
		/// <returns></returns>
		private string ParseLineForId(string line)
		{
			// this outer if is just to save some performance time
			if (line.IndexOf("[") > -1)
			{
				foreach (string key in ids.Keys)
				{
					string label = String.Format("[{0}]", key);
					line = line.Replace(label, ids[key].ToString());
				}
			}
			return line;
		}
		
		/// <summary>
		/// For children like sendePub will store in a hash for use when generating ePub files
		/// </summary>
		/// <param name="id"></param>
		/// <param name="idtext"></param>
		protected override void AddId(string id, string idtext)
		{
			if (ids.ContainsKey(id) == false)
			{
				ids.Add(id, idtext);
			}
		}
		public override string ToString ()
		{
			return string.Format ("[sendePub]");
		}
	}
}
