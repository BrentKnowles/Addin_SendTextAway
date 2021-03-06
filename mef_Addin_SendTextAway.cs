// mef_Addin_SendTextAway.cs
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
namespace MefAddIns
{

	using MefAddIns.Extensibility;
	using CoreUtilities;
	using System.ComponentModel.Composition;
	using System.ComponentModel;
	using System;
	using System.Windows.Forms;
	using System.Diagnostics;
	using System.Collections.Generic;
	using SendTextAway;
	using Layout;
	using System.IO;
	/// <summary>
	/// Provides an implementation of a supported language by implementing ISupportedLanguage. 
	/// Moreover it uses Export attribute to make it available thru MEF framework.
	/// </summary>
	[Export(typeof(mef_IBase))]
	public class Addin_SendTextAway :PlugInBase, mef_IBase
	{



		public Addin_SendTextAway()
		{
			guid = "sendtextaway";
		}
		
		public string Author
		{
			get { return @"Brent Knowles"; }
		}
		public string Version
		{
			// Version History
			// 1.1.2 - adding custom scene images and ellipsis
			// 1.1.1 - adding Chapter titles (if possible)
			// 1.1     -  getting output to produce 0 errors in "A Real Validator"
			// 1.0.8.2 - fixing [[fact]] errors, causing extra line feeds in epub version
			// 1.0.8.1 - handling &quot and a couple other HTML escape &lt &gt 
			// 1.0.8.0 - allow users to write [[~breakcomment]] to put a comment in that Calibre can do a find-replace on, to insert page breaks for PDF
			// 1.0.7.0 - allow links to other files in epub -- right now you can only go web or anchors inline
			// 1.0.6.0 - Attempting to handle multilevel (past 3) generically
			// 1.0.5.0 - handling 3rd level bullet points and backing out properly
			// 1.0.4.0 - trying to make Novels appear with lines indented (instead of flush-left). 01/05/2014
			// 1.0.3.0 - remove UNDERSCORES from Table of COntents -- what we show
			         //- Tweaked toc blank in TEMPLATE files for epub Sendaway to not have a title page.
			// 1.0.2.0 - series of bug fixes
			// 1.0.1.0 - adding <game> tagging
			get { return @"1.1.2.0"; }
		}
		public string Description
		{
			get { return "Exports the current note and converts it to formated text."; }
		}
		public string Name
		{
			get { return @"Send Text Away"; }
		}

		public static void GenerateFile(object param)
		{
			if (!(LayoutDetails.Instance.CurrentLayout.CurrentTextNote is NoteDataXML_SendIndex)) {
				NewMessage.Show (Loc.Instance.GetString ("You must use an Index file to send text to the sendaway system, even for individual notes."));
				return;
			}
			
			if (LayoutDetails.Instance.CurrentLayout != null && LayoutDetails.Instance.CurrentLayout.CurrentTextNote != null 
			    ) {
				
				sendBase SendAwayIt = null;
				
				ControlFile.convertertype TypeOfConverted = ((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.ConverterType;
				
				if (TypeOfConverted == ControlFile.convertertype.word)
				{
					SendAwayIt = new sendWord();
				}
				else if (TypeOfConverted == ControlFile.convertertype.epub)
				{
					SendAwayIt = new sendePub2();
				}
				else if (TypeOfConverted == ControlFile.convertertype.text)
				{
					SendAwayIt = new sendPlainText();
				}
				
				
				// error correction
				if (Constants.BLANK==((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.OutputDirectory ||
				    null == ((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.OutputDirectory)
				{
					string outputpath = Path.Combine(LayoutDetails.Instance.Path, "sendawayoutput");
					((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.OutputDirectory = outputpath;
				}
				
				if (Directory.Exists (((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.OutputDirectory ) == false)
				{
					Directory.CreateDirectory(((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.OutputDirectory );
				}
				
				//				ControlFile Control = new ControlFile ();
				//
				//				Control = (ControlFile)FileUtils.DeSerialize (@"C:\Users\BrentK\Documents\Keeper\SendTextAwayControlFiles\standardsub.xml", typeof(ControlFile));
				//				if (Control != null) {
				//NewMessage.Show ("Convering with " + ((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller.ConverterType.ToString());
				string sErrors = SendAwayIt.WriteText (param.ToString (), ((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller, -1);
				NewMessage.Show (sErrors);
				//}
				// will be used by this one
				//NewMessage.Show("SendAway " + param.ToString());
			} else {
				NewMessage.Show (Loc.Instance.GetString ("Please select a text note before using Send Text Away."));
			}
		}

		public void ActionWithParamForNoteTextActions (object param)
		{
			GenerateFile(param);
		}
		
		public void RespondToMenuOrHotkey<T>(T form) where T: System.Windows.Forms.Form, MEF_Interfaces.iAccess 
		{
			

		}
		public override bool DeregisterType ()
		{
			

			return true;

		}

		public override void RegisterType()
		{

			Layout.LayoutDetails.Instance.AddToList(typeof(NoteDataXML_SendIndex), Loc.Instance.GetStringFmt("Send Away Index"),
			                                        PlugInBase.AddInFolderName);
		}
		public override string dependencyguid {
			get {
				return "yourothermindmarkup";
			}
		}

		public static string BuildFileName()
		{
			return System.IO.Path.Combine (System.IO.Path.GetTempPath (), Guid.NewGuid().ToString () + ".txt");
		}

		public override string BuildFileNameForActionWithParam ()
		{
			return  BuildFileName();
		}
		public PlugInAction CalledFrom { 
			get
			{
				PlugInAction action = new PlugInAction();
		//		action.HotkeyNumber = -1;
				action.MyMenuName = String.Format ("{0} ({1})","Send Text Away", "Requires Valid Markup");// LayoutDetails.Instance.GetCurrentMarkup().ToString ());
				action.ParentMenuName = "";
				action.IsOnContextStrip = false;
				action.IsOnAMenu = false;

				// this is technically a note action but because I added the Generate button,
				// I don't see the purpose of having two routes to this.
				action.IsNoteAction = false;

				action.QuickLinkShows = false;
				action.IsANote = false;
				action.GUID = GUID;
				
				//action.IsOnAToolbar = false;
				return action;
			} 
		}
		
		
		
	}
}
