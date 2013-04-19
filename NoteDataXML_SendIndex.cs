// NoteDataXML_SendIndex.cs
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
using CoreUtilities;
using System.Windows.Forms;
using System.Drawing;
using CoreUtilities.Links;
using Layout;
using System.Xml.Serialization;
using SendTextAway;
namespace MefAddIns
{
	public class NoteDataXML_SendIndex  : Layout.NoteDataXML_RichText
	{
		public override int defaultHeight { get { return 500; } }
		public override int defaultWidth { get { return 300; } }
		#region variables
		public override bool IsLinkable { get { return false; }}
		

		ControlFile controller = ControlFile.Default;
		
		public ControlFile Controller {
			get {
				return controller;
			}
			set {
				controller = value;
			}
		}

		public bool UnderscoresUnderline {
			get { return !Controller.UnderlineShouldBeItalicInstead;}
			set { Controller.UnderlineShouldBeItalicInstead = !value;}
		}
		public ControlFile.convertertype ConverterType
		{
			get { return Controller.ConverterType;}
			set { Controller.ConverterType = value;}
		}

		public string BodyText
		{
			get { return Controller.BodyText;}
			set { Controller.BodyText = value;}
		}

		public bool UnderscoresRule
		{
			get { return Controller.UnderscoreKeep;}
			set { Controller.UnderscoreKeep = value;}
		}
	
		// the word template
		public string TemplateText
		{
			get { return Controller.Template;}
			set { Controller.Template = value;}
		}

#endregion
		
		#region interface
		TableLayoutPanel TablePanel = null;
	
#endregion

	
		public override void Dispose ()
		{
			
			
			base.Dispose();
			
		}
		private void CommonConstructor ()
		{
			Caption = Loc.Instance.GetString("Send Away Index");
		}
		public NoteDataXML_SendIndex () : base()
		{
			CommonConstructor();
		}
		public NoteDataXML_SendIndex(int height, int width):base(height, width)
		{
			CommonConstructor();
		}
		
		protected override void DoBuildChildren (LayoutPanelBase Layout)
		{
			base.DoBuildChildren (Layout);
			
			
			
			CaptionLabel.Dock = DockStyle.Top;
			
			TablePanel = new TableLayoutPanel ();



			TablePanel.Height = 200;
			TablePanel.RowCount = 5;
			TablePanel.ColumnCount = 2;
			TablePanel.Dock = DockStyle.Top;
			ParentNotePanel.Controls.Add (TablePanel);
			TablePanel.BringToFront ();
			//TablePanel.AutoSize = true;



		
			ToolTip Tipster = new ToolTip ();


			Label TypeOfSend = new Label ();
			TypeOfSend.Text = Loc.Instance.GetString ("Type");

			ComboBox ComboTypeOfSend = new ComboBox ();
			ComboTypeOfSend.DropDownStyle = ComboBoxStyle.DropDownList;
			//ComboTypeOfSend.DataSource = Enum.GetValues(typeof(ControlFile.convertertype));
			int selected = 0;
			ControlFile.convertertype[] vals = (ControlFile.convertertype[])Enum.GetValues (typeof(ControlFile.convertertype));
			for (int i = 0; i < vals.Length; i++) {
				ComboTypeOfSend.Items.Add (vals [i]);
				if (vals [i].ToString () == ConverterType.ToString ()) {
					selected = i;
				}
			}
			ComboTypeOfSend.SelectedIndex = selected;
			//NewMessage.Show ("Trying to set " + ConverterType.ToString ());
			//ComboTypeOfSend.SelectedText = ConverterType.ToString();
			//ComboTypeOfSend.SelectedValue = ConverterType.ToString();
			//ComboTypeOfSend.DataBindings();
			//ComboTypeOfSend.SelectedIndex = (int)ConverterType;
			ComboTypeOfSend.SelectedIndexChanged += HandleSelectedConverterTypeIndexChanged;


			Label Underscores = new Label ();
			Underscores.Dock = DockStyle.Fill;
			Underscores.Text = Loc.Instance.GetString ("Underscores...");

			CheckBox UnderscoresAsUnderline = new CheckBox ();
			UnderscoresAsUnderline.Dock = DockStyle.Top;
			UnderscoresAsUnderline.Width = 300;
			UnderscoresAsUnderline.Text = "Become Underline";
			UnderscoresAsUnderline.Checked = UnderscoresUnderline;
			Tipster.SetToolTip (UnderscoresAsUnderline, "If true underscore text will be underlined otherwise it will be in italics.");


			//Underscores.AutoSize = true;
			CheckBox UnderscoreKeep = new CheckBox ();
			UnderscoreKeep.Dock = DockStyle.Top;
			UnderscoreKeep.Text = Loc.Instance.GetString ("Keep Underscores");
			UnderscoreKeep.Checked = UnderscoresRule;
			Tipster.SetToolTip (UnderscoreKeep, Loc.Instance.GetString ("If set to true then underscores will always show up as underscore, no matter the other settings."));
			UnderscoreKeep.Click += (object sender, EventArgs e) => UnderscoresRule = (sender as CheckBox).Checked;

			UnderscoresAsUnderline.Click += (object sender, EventArgs e) => UnderscoresUnderline = (sender as CheckBox).Checked;


			Label BodyTextLabel = new Label ();
			BodyTextLabel.Dock = DockStyle.Fill;
			BodyTextLabel.Text = Loc.Instance.GetString ("Body Text");

			TextBox BodyTextText = new TextBox ();
			BodyTextText.Text = BodyText;
			BodyTextText.Dock = DockStyle.Fill;
			BodyTextText.TextChanged += (object sender, EventArgs e) => BodyText = (sender as TextBox).Text;
			BodyTextText.Width = 200; 

			Label TemplateLabel = new Label ();
			TemplateLabel.Dock = DockStyle.Fill;
			TemplateLabel.Text = Loc.Instance.GetString ("Template");

			TextBox TemplateTextBox = new TextBox ();
			TemplateTextBox.Text = TemplateText;
			TemplateTextBox.Dock = DockStyle.Fill;
			Tipster.SetToolTip (TemplateTextBox, Loc.Instance.GetString ("This is the Word template file, if generating a word document this will be the template used"));
			TemplateTextBox.TextChanged += (object sender, EventArgs e) => TemplateText = (sender as TextBox).Text;
			TemplateTextBox.Width = 200;

			// invokes a modal PropertyGrid for editing the entire fille
			Button EditAll = new Button ();
			EditAll.Text = Loc.Instance.GetString ("Edit All Details");
			EditAll.Click += HandleEditAllClick;
			EditAll.Dock = DockStyle.Fill;

			TablePanel.Controls.Add (TypeOfSend, 0, 0);
			TablePanel.Controls.Add (ComboTypeOfSend, 1, 0);

			TablePanel.Controls.Add (Underscores, 0, 1);
			TablePanel.Controls.Add (UnderscoresAsUnderline, 1, 1);
			TablePanel.Controls.Add (UnderscoreKeep, 1, 2);

			TablePanel.Controls.Add (BodyTextLabel, 0, 3);
			TablePanel.Controls.Add (BodyTextText, 1, 3);

			TablePanel.Controls.Add (TemplateLabel, 0, 4);
			TablePanel.Controls.Add (TemplateTextBox, 1, 4);


			Button Generate = new Button();
			Generate.Dock = DockStyle.Fill;
			Generate.Click+= HandleGenerateClick;
			Generate.Text = Loc.Instance.GetString ("Generate");

			TablePanel.Controls.Add (Generate, 0, 5);


			TablePanel.Controls.Add (EditAll, 1, 5);

//			TablePanel.ColumnStyles[0].SizeType  = SizeType.Percent;;
//			TablePanel.ColumnStyles[0].Width = 25;
//
//			TablePanel.ColumnStyles[1].SizeType  = SizeType.Percent;;
//			TablePanel.ColumnStyles[1].Width = 75;
//			foreach (ColumnStyle style in TablePanel.ColumnStyles) {
//			//	NewMessage.Show (style.ToString());
//				style.SizeType = SizeType.Percent;
//				style.Width = 50;
//			}

			if (richBox.Text == Constants.BLANK) {
				richBox.Text = Loc.Instance.GetStringFmt("[[index]]{0}Enter Page Name Here Followed By Line Space{0}", Environment.NewLine);
			}
			richBox.BringToFront();
			
		}

		/// <summary>
		/// A shortcut button to 'send away'
		/// </summary>
		/// <param name='sender'>
		/// Sender.
		/// </param>
		/// <param name='e'>
		/// E.
		/// </param>
		void HandleGenerateClick (object sender, EventArgs e)
		{

			if (LayoutDetails.Instance.CurrentLayout != null) {
				// this is tricky because the noteaction operation took care of writing the file so we need to repeat this.
				LayoutDetails.Instance.CurrentLayout.CurrentTextNote = this;

				string FileToSaveTo = MefAddIns.Addin_SendTextAway.BuildFileName ();


				string[] lines = LayoutDetails.Instance.CurrentLayout.CurrentTextNote.Lines ();
				if (lines.Length > 0) {
					LayoutDetails.Instance.SaveTextLineToFile (lines, FileToSaveTo);
					// now process  saved text
					MefAddIns.Addin_SendTextAway.GenerateFile (FileToSaveTo);
				}


			}
		}

		void HandleEditAllClick (object sender, EventArgs e)
		{
			ControlFilePropertyEditForm form = new ControlFilePropertyEditForm(this.Controller);
			form.ShowDialog();
		}

		void HandleSelectedConverterTypeIndexChanged (object sender, EventArgs e)
		{
			if ((sender as ComboBox).SelectedItem != null) {
				ControlFile.convertertype conv = ControlFile.convertertype.text;
				Enum.TryParse<ControlFile.convertertype>((sender as ComboBox).SelectedItem.ToString (), out conv);
				this.Controller.ConverterType = conv;
			}
		}
		

		protected override void DoChildAppearance (AppearanceClass app)
		{
			base.DoChildAppearance (app);
			
			TablePanel.BackColor = app.mainBackground;
			
		}
		public override void Save ()
		{
			base.Save ();
			//CharacterColorInt = CharacterColor.ToArgb();
		}
		
		
		/// <summary>
		/// Registers the type.
		/// </summary>
		public override string RegisterType()
		{
			return Loc.Instance.GetString("Send Away Index");
		}

		public NoteDataXML_SendIndex(NoteDataInterface Note) : base(Note)
		{
			
		}
		public override void CopyNote (NoteDataInterface Note)
		{
			base.CopyNote (Note);
			if (Note is NoteDataXML_SendIndex) {
				this.CopyObject ((Note as NoteDataXML_SendIndex).Controller, this.Controller);

			}
		}
		
	}
}

