namespace MefAddIns
{

	using MefAddIns.Extensibility;
	using CoreUtilities;
	using System.ComponentModel.Composition;
	using System;
	using System.Windows.Forms;
	using System.Diagnostics;
	using System.Collections.Generic;
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
		
		public void ActionWithParam (object param)
		{
			// will be used by this one
			NewMessage.Show("SendAway " + param.ToString());
		}
		
		public void RespondToCallToAction ()
		{
			
		
		}
		public override string BuildFileName ()
		{
			return  System.IO.Path.Combine (System.IO.Path.GetTempPath (), Guid.NewGuid().ToString () + ".txt");
		}
		public PlugInAction CalledFrom { 
			get
			{
				PlugInAction action = new PlugInAction();
				action.HotkeyNumber = -1;
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
