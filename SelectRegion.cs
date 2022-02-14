using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BruteClipper
{
    public partial class SelectRegion : Form
    {
        FormMask _formMaskTop;
        FormMask _formMaskLeft;
        FormMask _formMaskRight;
        FormMask _formMaskBottom;

        public SelectRegion(Rectangle startingPosition)
        {
            InitializeComponent();

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(startingPosition.Left, startingPosition.Top);
            this.Size = new Size(startingPosition.Width, startingPosition.Height);

            this.Move += new EventHandler(MoveSubForm);
            this.Resize += new EventHandler(MoveSubForm);
        }

        private void SelectRegion_Load(object sender, EventArgs e)
        {
            _formMaskTop = new FormMask();
            _formMaskLeft = new FormMask();
            _formMaskRight = new FormMask();
            _formMaskBottom = new FormMask();

            _formMaskTop.Click += new EventHandler(SelectRegion_Click);
            _formMaskLeft.Click += new EventHandler(SelectRegion_Click);
            _formMaskRight.Click += new EventHandler(SelectRegion_Click);
            _formMaskBottom.Click += new EventHandler(SelectRegion_Click);

            _formMaskTop.Show();
            _formMaskLeft.Show();
            _formMaskRight.Show();
            _formMaskBottom.Show();

            MoveSubForm(this, e);
        }

        protected void MoveSubForm(object sender, EventArgs e)
        {
            var screenBounds = Screen.GetBounds(Point.Empty);

            if (_formMaskTop != null)
            {
                _formMaskTop.Top = 0;
                _formMaskTop.Left = 0;
                _formMaskTop.Width = screenBounds.Width;
                _formMaskTop.Height = this.Top;
            }
            if (_formMaskLeft != null)
            {
                _formMaskLeft.Top = this.Top;
                _formMaskLeft.Left = 0;
                _formMaskLeft.Width = this.Left;
                _formMaskLeft.Height = this.Height;
            }
            if (_formMaskRight != null)
            {
                _formMaskRight.Top = this.Top;
                _formMaskRight.Left = this.Right;
                _formMaskRight.Width = screenBounds.Width - this.Right;
                _formMaskRight.Height = this.Height;
            }
            if (_formMaskBottom != null)
            {
                _formMaskBottom.Top = this.Bottom;
                _formMaskBottom.Left = 0;
                _formMaskBottom.Width = screenBounds.Width;
                _formMaskBottom.Height = screenBounds.Height - this.Bottom;
            }
        }

        private void SelectRegion_FormClosing(object sender, FormClosingEventArgs e)
        {
            _formMaskTop.Dispose();
            _formMaskLeft.Dispose();
            _formMaskRight.Dispose();
            _formMaskBottom.Dispose();
        }



        private const int cGrip = 16;      // Grip size
        //private const int cCaption = 32;   // Caption bar height;

        private void SelectRegion_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rc = new Rectangle(this.ClientSize.Width - cGrip, this.ClientSize.Height - cGrip, cGrip, cGrip);
            ControlPaint.DrawSizeGrip(e.Graphics, this.BackColor, rc);
            //rc = new Rectangle(0, 0, this.ClientSize.Width, cCaption);
            //e.Graphics.FillRectangle(Brushes.DarkBlue, rc);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x84)
            {  // Trap WM_NCHITTEST
                Point pos = new Point(m.LParam.ToInt32());
                pos = this.PointToClient(pos);
                //if (pos.Y < cCaption)
                //{
                //    m.Result = (IntPtr)2;  // HTCAPTION
                //    return;
                //}
                if (pos.X >= this.ClientSize.Width - cGrip && pos.Y >= this.ClientSize.Height - cGrip)
                {
                    m.Result = (IntPtr)17; // HTBOTTOMRIGHT
                    return;
                }
            }
            base.WndProc(ref m);
        }



        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void SelectRegion_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }



        private void SelectRegion_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            FrmMain.SelectRegionResult = new Rectangle(this.Left, this.Top, this.Width, this.Height);
            this.Dispose();
        }

        private void SelectRegion_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Dispose();
            }
        }
    }

    public partial class FormMask : Form
    {
        public FormMask()
        {
            this.BackColor = Color.Black;
            this.Opacity = 0.5;
            this.ShowInTaskbar = false;
            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.TopMost = true;
        }
    }


}
