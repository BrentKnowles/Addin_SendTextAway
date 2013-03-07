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
			get { return @"1.0.0.1"; }
		}
		public string Description
		{
			get { return "Exports the current note and converts it to formated text."; }
		}
		public string Name
		{
			get { return @"Send Text Away"; }
		}
		
		public void ActionWithParamForNoteTextActions (object param)
		{

			if (LayoutDetails.Instance.CurrentLayout != null && LayoutDetails.Instance.CurrentLayout.CurrentTextNote != null 
			    && LayoutDetails.Instance.CurrentLayout.CurrentTextNote is NoteDataXML_SendIndex) {

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
				SendAwayIt.WriteText (param.ToString (), ((NoteDataXML_SendIndex) LayoutDetails.Instance.CurrentLayout.CurrentTextNote).Controller, -1);
			
				//}
				// will be used by this one
				//NewMessage.Show("SendAway " + param.ToString());
			} else {
				NewMessage.Show (Loc.Instance.GetString ("Please select a text note before using Send Text Away"));
			}
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

			Layout.LayoutDetails.Instance.AddToList(typeof(NoteDataXML_SendIndex), Loc.Instance.GetString ("Send Away Index"));
		}
		public override string dependencyguid {
			get {
				//TODO remove. This does not need a dependency. Just testing
				return "yourothermindmarkup";
			}
		}

		public override string BuildFileNameForActionWithParam ()
		{
			return  System.IO.Path.Combine (System.IO.Path.GetTempPath (), Guid.NewGuid().ToString () + ".txt");
		}
		public PlugInAction CalledFrom { 
			get
			{
				PlugInAction action = new PlugInAction();
		//		action.HotkeyNumber = -1;
				action.MyMenuName = "Send Text Away";
				action.ParentMenuName = "";
				action.IsOnContextStrip = false;
				action.IsOnAMenu = false;
				action.IsNoteAction = true;

				action.QuickLinkShows = false;
				action.IsANote = false;
				action.GUID = GUID;
				
				//action.IsOnAToolbar = false;
				return action;
			} 
		}
		
		
		
	}
}
