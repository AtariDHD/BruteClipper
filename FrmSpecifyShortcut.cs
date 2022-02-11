using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace BruteClipper
{
	/// <summary>
	/// Summary description for FrmSpecifyShortcut.
	/// </summary>
	public class FrmSpecifyShortcut : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button BtnOK;
		private System.Windows.Forms.Button BtnCancel;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public BruteClipper.ShortcutInput ShortcutInput1;
		private byte MinModifiers;


		public FrmSpecifyShortcut(byte minMod, Keys k)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			MinModifiers = minMod;
			ShortcutInput1.MinModifiers = MinModifiers;

			if (k != Keys.None)
                ShortcutInput1.Keys = k;
		}


		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.BtnOK = new System.Windows.Forms.Button();
			this.BtnCancel = new System.Windows.Forms.Button();
			this.ShortcutInput1 = new BruteClipper.ShortcutInput();
			this.SuspendLayout();
			// 
			// BtnOK
			// 
			this.BtnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.BtnOK.Location = new System.Drawing.Point(48, 88);
			this.BtnOK.Name = "BtnOK";
			this.BtnOK.TabIndex = 0;
			this.BtnOK.Text = "&OK";
			this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
			// 
			// BtnCancel
			// 
			this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.BtnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.BtnCancel.Location = new System.Drawing.Point(128, 88);
			this.BtnCancel.Name = "BtnCancel";
			this.BtnCancel.TabIndex = 1;
			this.BtnCancel.Text = "&Cancel";
			// 
			// ShortcutInput1
			// 
			this.ShortcutInput1.Alt = false;
			this.ShortcutInput1.CharCode = ((System.Byte)(65));
			this.ShortcutInput1.Control = false;
			this.ShortcutInput1.Keys = System.Windows.Forms.Keys.A;
			this.ShortcutInput1.Location = new System.Drawing.Point(8, 16);
			this.ShortcutInput1.Name = "ShortcutInput1";
			this.ShortcutInput1.Shift = false;
			this.ShortcutInput1.Size = new System.Drawing.Size(248, 48);
			this.ShortcutInput1.TabIndex = 2;
			// 
			// FrmSpecifyShortcut
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.BtnCancel;
			this.ClientSize = new System.Drawing.Size(250, 122);
			this.Controls.Add(this.ShortcutInput1);
			this.Controls.Add(this.BtnCancel);
			this.Controls.Add(this.BtnOK);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "FrmSpecifyShortcut";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Specify Shortcut";
			this.ResumeLayout(false);

		}
		#endregion

		
		private void BtnOK_Click(object sender, System.EventArgs e)
		{
			if (!ShortcutInput1.IsValid)
			{
				MessageBox.Show("Please check at least " + MinModifiers + " of the checkboxes (ie Modifier keys)   ", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			DialogResult = DialogResult.OK;
			Close();
		}

	}
}
