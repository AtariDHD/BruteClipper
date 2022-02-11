using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;


namespace BruteClipper
{
    /// <summary>
    /// Summary description for FrmMain.
    /// </summary>
    public class FrmMain : System.Windows.Forms.Form
	{
		private System.Windows.Forms.NumericUpDown NumMinMod;
		private System.Windows.Forms.TextBox TxtKeyEnumValue;
		private System.Windows.Forms.Button BtnChangeHotkey;
        private NotifyIcon SystemTrayIcon;
        private System.ComponentModel.IContainer components;

        private bool allowVisible;     // ContextMenu's Open command used

        public FrmMain()
		{
			InitializeComponent();

			this.CenterToScreen();

            // To provide your own custom icon image, go to:
            //   1. Project > Properties... > Resources
            //   2. Change the resource filter to icons
            //   3. Remove the Default resource and add your own
            //   4. Modify the next lines to Properties.Resources.<YourResource>
            this.Icon = Properties.Resources.Default;
			this.SystemTrayIcon.Icon = Properties.Resources.Default;

			this.SystemTrayIcon.Text = "Brute Clipper";
			this.SystemTrayIcon.Visible = true;

			ContextMenu menu = new ContextMenu();
            menu.MenuItems.Add(new MenuItem { Text = "Brute Clipper", Enabled = false });
            menu.MenuItems.Add("Open", ContextMenuOpen);
			menu.MenuItems.Add("Exit", ContextMenuExit);
			this.SystemTrayIcon.ContextMenu = menu;

            this.Resize += WindowResize;
			this.FormClosing += WindowClosing;

            LoadHotKeys();
        }

        protected override void SetVisibleCore(bool value)
        {
            if (!allowVisible)
            {
                value = false;
                if (!this.IsHandleCreated) CreateHandle();
            }
            base.SetVisibleCore(value);
        }

        private void SystemTrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            OpenAppWindow();
        }

        private void ContextMenuOpen(object sender, EventArgs e)
        {
            OpenAppWindow();
        }

        private void ContextMenuExit(object sender, EventArgs e)
        {
            FrmMain.UnregisterHotKey(this.Handle, this.GetType().GetHashCode());
            this.SystemTrayIcon.Visible = false;
            Application.Exit();
            Environment.Exit(0);
        }

        private void OpenAppWindow()
        {
            allowVisible = true;
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void WindowResize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
            }
        }

        private void WindowClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
        }

        private void LoadHotKeys()
        {
            if (File.Exists(Application.StartupPath + "\\HotkeyValue.txt"))
            {
                StreamReader reader = File.OpenText(Application.StartupPath + "\\HotkeyValue.txt");
                int val = -1;
                try
                {
                    val = Int32.Parse(reader.ReadToEnd().Trim());
                }
                catch { }
                if (val != -1)
                {
                    Keys k = (Keys)val;
                    bool success = FrmMain.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), ShortcutInput.Win32ModifiersFromKeys(k), ShortcutInput.CharCodeFromKeys(k));
                    if (success)
                        TxtKeyEnumValue.Text = val.ToString();
                    else
                        MessageBox.Show("Could not register Hotkey - there is probably a conflict.  ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.NumMinMod = new System.Windows.Forms.NumericUpDown();
            this.TxtKeyEnumValue = new System.Windows.Forms.TextBox();
            this.BtnChangeHotkey = new System.Windows.Forms.Button();
            this.SystemTrayIcon = new System.Windows.Forms.NotifyIcon(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.NumMinMod)).BeginInit();
            this.SuspendLayout();
            // 
            // NumMinMod
            // 
            this.NumMinMod.Location = new System.Drawing.Point(-2, 77);
            this.NumMinMod.Maximum = new decimal(new int[] {
            3,
            0,
            0,
            0});
            this.NumMinMod.Name = "NumMinMod";
            this.NumMinMod.Size = new System.Drawing.Size(48, 20);
            this.NumMinMod.TabIndex = 1;
            this.NumMinMod.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.NumMinMod.Visible = false;
            // 
            // TxtKeyEnumValue
            // 
            this.TxtKeyEnumValue.Location = new System.Drawing.Point(52, 77);
            this.TxtKeyEnumValue.Name = "TxtKeyEnumValue";
            this.TxtKeyEnumValue.ReadOnly = true;
            this.TxtKeyEnumValue.Size = new System.Drawing.Size(88, 20);
            this.TxtKeyEnumValue.TabIndex = 3;
            this.TxtKeyEnumValue.Visible = false;
            // 
            // BtnChangeHotkey
            // 
            this.BtnChangeHotkey.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnChangeHotkey.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnChangeHotkey.Location = new System.Drawing.Point(44, 26);
            this.BtnChangeHotkey.Name = "BtnChangeHotkey";
            this.BtnChangeHotkey.Size = new System.Drawing.Size(193, 45);
            this.BtnChangeHotkey.TabIndex = 4;
            this.BtnChangeHotkey.Text = "Change Paste Hotkey";
            this.BtnChangeHotkey.Click += new System.EventHandler(this.BtnChangeHotkey_Click);
            // 
            // SystemTrayIcon
            // 
            this.SystemTrayIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.SystemTrayIcon.Visible = true;
            this.SystemTrayIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(this.SystemTrayIcon_MouseClick);
            // 
            // FrmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(277, 98);
            this.Controls.Add(this.BtnChangeHotkey);
            this.Controls.Add(this.TxtKeyEnumValue);
            this.Controls.Add(this.NumMinMod);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(293, 137);
            this.MinimumSize = new System.Drawing.Size(293, 137);
            this.Name = "FrmMain";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Brute Clipper";
            ((System.ComponentModel.ISupportInitialize)(this.NumMinMod)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

	
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new FrmMain());
		}


		private void BtnChangeHotkey_Click(object sender, System.EventArgs e)
		{
			byte minMod = (byte)NumMinMod.Value;
			Keys k = (TxtKeyEnumValue.Text.Length > 0) ? (Keys)Int32.Parse(TxtKeyEnumValue.Text) : Keys.None;
			FrmSpecifyShortcut frm = new FrmSpecifyShortcut(minMod, k);
			if (frm.ShowDialog() == DialogResult.OK)
			{
                FrmMain.UnregisterHotKey(this.Handle, this.GetType().GetHashCode());
                Thread.Sleep(100);
                bool success = FrmMain.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), frm.ShortcutInput1.Win32Modifiers, frm.ShortcutInput1.CharCode);
				if (success)
				{
					TxtKeyEnumValue.Text = ((int)frm.ShortcutInput1.Keys).ToString();
					StreamWriter writer = File.CreateText(Application.StartupPath + "\\HotkeyValue.txt");
					writer.Write(TxtKeyEnumValue.Text);
					writer.Close();
				}
				else
					MessageBox.Show("Brute Clipper: Could not register hotkey - there is probably a conflict.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


        [DllImport("user32.dll")]
		public static extern bool RegisterHotKey(IntPtr hWnd,int id,int fsModifiers,int vlc);
		[DllImport("user32.dll")]
		public static extern bool UnregisterHotKey(IntPtr hWnd, int id);


		protected override void WndProc(ref Message m)
		{
			if (m.Msg == 0x0312)
            {
                if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
                {
                    Thread.Sleep(1200); // TODO: need a better way of making sure user is still not holding down hotkey modifier keys
                    var keys = EscapeSendKeysSpecialCharacters(Clipboard.GetText(TextDataFormat.UnicodeText));
                    System.Diagnostics.Debug.WriteLine("pasting unicode");
                    SendKeys.SendWait(keys);
                }
                else if (Clipboard.ContainsText(TextDataFormat.Text))
                {
                    Thread.Sleep(1200); // TODO: need a better way of making sure user is still not holding down hotkey modifier keys
                    var keys = EscapeSendKeysSpecialCharacters(Clipboard.GetText(TextDataFormat.Text));
                    System.Diagnostics.Debug.WriteLine("pasting text");
                    SendKeys.SendWait(keys);
                }
                else
                {
                    //MessageBox.Show("No clipboard Text or UnicodeText");
                }
            }
            base.WndProc(ref m);
		}

        private string EscapeSendKeysSpecialCharacters(string str)
        {
            var reSendKeysChars = new Regex(@"(?<SpecialCharacter>[+^%~{}[\]])");
            var escaped = reSendKeysChars.Replace(str, m => m.Value.Replace(m.Groups["SpecialCharacter"].Value, $"{{{m.Groups["SpecialCharacter"].Value}}}"));

            escaped = escaped.Replace("\r\n", "\n");

            return escaped;
        }
    }
}
