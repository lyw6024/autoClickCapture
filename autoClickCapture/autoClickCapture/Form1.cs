using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using OpenCvSharp;

namespace autoClickCapture
{
    public partial class Form1 : Form
    {

        BackgroundWorker bgClick;
        public Form1()
        {
            InitializeComponent();
            bgClick = new BackgroundWorker();
            bgClick.DoWork += new DoWorkEventHandler(clickbgwf);
            bgClick.WorkerReportsProgress = true;
            bgClick.ProgressChanged += new ProgressChangedEventHandler(clickpcbgwf);
            bgClick.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
                (object s,RunWorkerCompletedEventArgs e) => {button1.Enabled=true; }) ;

        }

        [DllImport("User32")]
        public extern static void mouse_event(int dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);

        public enum MouseEventFlags       //鼠标按键的ASCLL码
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            Wheel = 0x0800,
            Absolute = 0x8000
        }

        private void AutoClick()
        {
            mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);//自动按下的按键
            mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
        }

        private void clickbgwf(object sender,DoWorkEventArgs e)
        {
            int endNumber = 0;
            if (e.Argument != null)
            {
                endNumber = (int)e.Argument;
            }
            for (int i=0;i< endNumber; i++)
            {
                Thread.Sleep((int)(numericUpDown2.Value - numericUpDown3.Value));
                AutoClick();
                Thread.Sleep((int)(numericUpDown3.Value));
                ((BackgroundWorker)sender).ReportProgress(0,i);
            }
        }

        Bitmap lastb;
        Image lastImg;
        Mat lastsrc,lastdst;
        private void clickpcbgwf(object sender,ProgressChangedEventArgs e)
        {
            if (lastImg != null)
            {
                lastImg.Dispose();
            }
            if(lastb!=null)
            {
                lastb.Dispose();
            }
            if(lastsrc!=null)
            {
                lastsrc.Dispose();
            }
            if(lastdst!=null)
            {
                lastdst.Dispose();
            }
            numericUpDown1.Value -= 1;
            Bitmap b = GetScreenCapture();
            Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(b);
            Mat dst = new Mat();
            src.ConvertTo(dst, MatType.CV_8UC1);
            dst.SaveImage(String.Format("shot/cap{0:0000}.png", (int)e.UserState));
            Image p = (Image)b;
            pictureBox1.BackgroundImage = p;

            lastImg = p;
            lastb = b;
            lastsrc = src;
            lastdst = dst;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            bgClick.RunWorkerAsync((int)numericUpDown1.Value);
        }

        private Bitmap GetScreenCapture()
        {
            Rectangle tScreenRect = new Rectangle(0, 0, 565, 1000);
            Bitmap tSrcBmp = new Bitmap(tScreenRect.Width, tScreenRect.Height); // 用于屏幕原始图片保存
            Graphics gp = Graphics.FromImage(tSrcBmp);
            gp.CopyFromScreen(0, 38, 0, 0, tScreenRect.Size);
            gp.DrawImage(tSrcBmp, 0, 0, tScreenRect, GraphicsUnit.Pixel);
            return tSrcBmp;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown3.Maximum = ((NumericUpDown)sender).Value;
        }
    }
}
