using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Tesseract;

namespace BruteClipper
{
    /// <summary>
    /// Summary description for FrmMain.
    /// </summary>
    public class FrmMain : Form
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public const int WM_HOTKEY = 0x0312;
        private const int COPY_HOTKEY_ID = 1;
        private const int PASTE_HOTKEY_ID = 2;

        private const int MinModifierKeys = 1;

        private ScreenCapturer screenCapturer = new ScreenCapturer();

        private Button BtnChangeCopyHotkey;
        private Button BtnChangePasteHotkey;
        private NotifyIcon SystemTrayIcon;
        private System.ComponentModel.IContainer components;

        private bool allowVisible; // ContextMenu's Open command used

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmMain());
        }

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

            LoadHotkeys();
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
            if (e.Button == MouseButtons.Left)
            {
                OpenAppWindow();
            }
        }

        private void ContextMenuOpen(object sender, EventArgs e)
        {
            OpenAppWindow();
        }

        private void ContextMenuExit(object sender, EventArgs e)
        {
            UnregisterHotKey(this.Handle, PASTE_HOTKEY_ID);
            SystemTrayIcon.Visible = false;
            Application.Exit();
            Environment.Exit(0);
        }

        private void OpenAppWindow()
        {
            allowVisible = true;
            this.Show();
            this.Activate();
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

        private void LoadHotkeys()
        {
            LoadHotkey("CopyHotkey", COPY_HOTKEY_ID);
            LoadHotkey("PasteHotkey", PASTE_HOTKEY_ID);
        }

        private void LoadHotkey(string SettingName, int HotkeyId)
        {
            if ((int)Properties.Settings.Default[SettingName] > 0)
            {
                Keys k = (Keys)Properties.Settings.Default[SettingName];
                bool success = RegisterHotKey(this.Handle, HotkeyId, ShortcutInput.Win32ModifiersFromKeys(k), ShortcutInput.CharCodeFromKeys(k));
                if (!success)
                    MessageBox.Show("Brute Clipper: Could not register hotkey. There is probably a conflict.  ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnChangeCopyHotkey_Click(object sender, EventArgs e)
        {
            SpecifyShortcut("CopyHotkey", COPY_HOTKEY_ID);
        }

        private void BtnChangePasteHotkey_Click(object sender, EventArgs e)
        {
            SpecifyShortcut("PasteHotkey", PASTE_HOTKEY_ID);
        }

        private void SpecifyShortcut(string SettingName, int HotkeyId)
        {
            Keys k = (int)Properties.Settings.Default[SettingName] > 0 ? (Keys)Properties.Settings.Default[SettingName] : Keys.None;

            FrmSpecifyShortcut frm = new FrmSpecifyShortcut(MinModifierKeys, k);

            if (frm.ShowDialog() == DialogResult.OK)
            {
                UnregisterHotKey(this.Handle, HotkeyId);
                Thread.Sleep(100);
                bool success = RegisterHotKey(this.Handle, HotkeyId, frm.ShortcutInput1.Win32Modifiers, frm.ShortcutInput1.CharCode);
                if (success)
                {
                    Properties.Settings.Default[SettingName] = (int)frm.ShortcutInput1.Keys;
                    Properties.Settings.Default.Save();
                }
                else
                    MessageBox.Show("Brute Clipper: Could not register hotkey. There is probably a conflict.  ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                if (m.WParam.ToInt32() == COPY_HOTKEY_ID)
                {
                    SystemTrayIcon.ShowBalloonTip(2000, "Brute Clipper", "Copy OCR Started", ToolTipIcon.Info);
                    var activeWindowScreenGrab = screenCapturer.Capture();
                    //Clipboard.SetImage(activeWindowScreenGrab);
                    BitmapOcrToClipboard(activeWindowScreenGrab);
                }
                else if (m.WParam.ToInt32() == PASTE_HOTKEY_ID)
                {
                    if (Clipboard.ContainsText(TextDataFormat.Text) || Clipboard.ContainsText(TextDataFormat.UnicodeText))
                    {
                        Thread.Sleep(1200); // TODO: need a better way of making sure user is still not holding down hotkey modifier keys
                        var keys = EscapeSendKeysSpecialCharacters(Clipboard.GetText());
                        Debug.WriteLine("Brute Clipper: Hotkey called. Pasting text.");
                        SendKeys.SendWait(keys);
                    }
                    else
                    {
                        Debug.WriteLine("Brute Clipper: Hotkey called, but no text in clipboard to paste.");
                    }
                }
            }
            base.WndProc(ref m);
        }

        private string BitmapOcrToClipboard(Bitmap image)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./Resources/tessdata", "eng", EngineMode.Default))
                {
                    var bmpToPix = new Tesseract.BitmapToPixConverter();

                    using (var img = bmpToPix.Convert(image))
                    {
                        using (var page = engine.Process(img))
                        {
                            var text = page.GetText();

                            Debug.WriteLine("Mean confidence: {0}", page.GetMeanConfidence());
                            //Debug.WriteLine(text);
                            Clipboard.SetText(text);
                            return text;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Trace.TraceError(e.ToString());
                Debug.WriteLine("Unexpected Error: " + e.Message);
                Debug.WriteLine("Details: ");
                Debug.WriteLine(e.ToString());
                return null;
            }
        }

        private string EscapeSendKeysSpecialCharacters(string str)
        {
            var reSendKeysChars = new Regex(@"(?<SpecialCharacter>[+^%~{}[\]\(\)])");
            var escaped = reSendKeysChars.Replace(str, m => m.Value.Replace(m.Groups["SpecialCharacter"].Value, $"{{{m.Groups["SpecialCharacter"].Value}}}"));
            escaped = escaped.Replace("\r\n", "~");

            return escaped;
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

        public enum ScreenCaptureMode
        {
            Screen,
            Window
        }

        class ScreenCapturer
        {
            [DllImport("user32.dll")]
            private static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

            [StructLayout(LayoutKind.Sequential)]
            private struct Rect
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            public Bitmap Capture(ScreenCaptureMode screenCaptureMode = ScreenCaptureMode.Window)
            {
                Rectangle bounds;

                if (screenCaptureMode == ScreenCaptureMode.Screen)
                {
                    bounds = Screen.GetBounds(Point.Empty);
                }
                else
                {
                    var foregroundWindowsHandle = GetForegroundWindow();
                    var rect = new Rect();
                    GetWindowRect(foregroundWindowsHandle, ref rect);
                    bounds = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
                }

                var result = new Bitmap(bounds.Width, bounds.Height);

                using (var g = Graphics.FromImage(result))
                {
                    g.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
                }

                return result;
            }
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.BtnChangePasteHotkey = new System.Windows.Forms.Button();
            this.SystemTrayIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.BtnChangeCopyHotkey = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // BtnChangePasteHotkey
            // 
            this.BtnChangePasteHotkey.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnChangePasteHotkey.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnChangePasteHotkey.Location = new System.Drawing.Point(43, 106);
            this.BtnChangePasteHotkey.Name = "BtnChangePasteHotkey";
            this.BtnChangePasteHotkey.Size = new System.Drawing.Size(231, 52);
            this.BtnChangePasteHotkey.TabIndex = 4;
            this.BtnChangePasteHotkey.Text = "Change Paste Hotkey";
            this.BtnChangePasteHotkey.Click += new System.EventHandler(this.BtnChangePasteHotkey_Click);
            // 
            // SystemTrayIcon
            // 
            this.SystemTrayIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.SystemTrayIcon.Visible = true;
            this.SystemTrayIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(this.SystemTrayIcon_MouseClick);
            // 
            // BtnChangeCopyHotkey
            // 
            this.BtnChangeCopyHotkey.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnChangeCopyHotkey.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnChangeCopyHotkey.Location = new System.Drawing.Point(43, 29);
            this.BtnChangeCopyHotkey.Name = "BtnChangeCopyHotkey";
            this.BtnChangeCopyHotkey.Size = new System.Drawing.Size(231, 53);
            this.BtnChangeCopyHotkey.TabIndex = 5;
            this.BtnChangeCopyHotkey.Text = "Change Copy Hotkey";
            this.BtnChangeCopyHotkey.Click += new System.EventHandler(this.BtnChangeCopyHotkey_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.ClientSize = new System.Drawing.Size(314, 186);
            this.Controls.Add(this.BtnChangeCopyHotkey);
            this.Controls.Add(this.BtnChangePasteHotkey);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FrmMain";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Brute Clipper";
            this.ResumeLayout(false);

        }
        #endregion
    }
}
