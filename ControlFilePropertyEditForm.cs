using System;
using System.Windows.Forms;
using Layout;
using CoreUtilities;

namespace SendTextAway
{
	public class ControlFilePropertyEditForm : Form
	{
		public ControlFilePropertyEditForm (ControlFile Control)
		{
			this.Height = 500;
			this.Width = 500;
			if (null == Control) {
				throw new Exception("Must specify a valid control file.");
			}
			this.Icon = LayoutDetails.Instance.MainFormIcon;
			FormUtils.SizeFormsForAccessibility(this, LayoutDetails.Instance.MainFormFontSize);

			Button OK = new Button();
			OK.Text = Loc.Instance.GetString ("OK");
			OK.Dock =  DockStyle.Bottom;
			OK.DialogResult = DialogResult.OK;
			this.Controls.Add (OK);


			PropertyGrid grid = new PropertyGrid();
			grid.Dock = DockStyle.Fill;
			this.Controls.Add (grid);
			grid.BringToFront();

			grid.SelectedObject = Control;



		}
	}
}

