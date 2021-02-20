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
using System.Data.SQLite;
using OpenCvSharp;

namespace autoClickCapture
{
    public partial class Form1 : Form
    {

        BackgroundWorker bgClick,oncount;
        SQLiteConnection m_dbConnection;

        public Form1()
        {
            InitializeComponent();
            bgClick = new BackgroundWorker();
            bgClick.WorkerReportsProgress = true;
            bgClick.WorkerSupportsCancellation = true;
            bgClick.DoWork += new DoWorkEventHandler(clickbgwf);
            bgClick.ProgressChanged += new ProgressChangedEventHandler(clickprogresscgbgwf);
            bgClick.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
                (object s,RunWorkerCompletedEventArgs e) => { setCapsValue(true);progressBar1.Value = 0; }) ;

            oncount = new BackgroundWorker();
            oncount.WorkerReportsProgress = true;
            oncount.DoWork += new DoWorkEventHandler(
                (s,e)=>
                {
                    for (int i = 0; i < 3; i++)
                    {
                        Thread.Sleep(1000);
                        ((BackgroundWorker)s).ReportProgress(i);
                    }
                });

            oncount.ProgressChanged += new ProgressChangedEventHandler(
                (s, e) =>{progressBar1.Value = 1+e.ProgressPercentage;});

            oncount.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
                (s, e) => { bgClick.RunWorkerAsync((int)numericUpDown1.Value); }) ;
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
            Thread.Sleep(1000);
            int endNumber = 0;
            if (e.Argument != null)
            {
                endNumber = (int)e.Argument;
            }
            for (int i=0;i< endNumber; i++)
            {
                if(bgClick.CancellationPending){break;}
                Thread.Sleep((int)(numericUpDown2.Value - numericUpDown3.Value));
                if (bgClick.CancellationPending) { break; }
                AutoClick();
                if (bgClick.CancellationPending) { break; }
                Thread.Sleep((int)(numericUpDown3.Value));
                if (bgClick.CancellationPending) { break; }
                ((BackgroundWorker)sender).ReportProgress(0,i);
                if (bgClick.CancellationPending) { break; }
            }
        }

        Bitmap lastb;
        Image lastImg;
        Mat lastsrc,lastdst;
        private void clickprogresscgbgwf(object sender,ProgressChangedEventArgs e)
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
            if(checkBox1.Checked)
            {
                Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(b);
                Mat dst = new Mat();
                src.ConvertTo(dst, MatType.CV_8UC1);
                dst.SaveImage(String.Format("shot/cap{0:0000}.png", (int)e.UserState));
                lastsrc = src;
                lastdst = dst;
            }
            Image p = (Image)b;
            pictureBox1.BackgroundImage = p;

            lastImg = p;
            lastb = b;
        }

        private void setCapsValue(bool s)
        {
            button1.Enabled = s;
            numericUpDown1.Enabled = s;
            numericUpDown2.Enabled = s;
            numericUpDown3.Enabled = s;
            leftUpperY.Enabled = s;
            leftUpperX.Enabled = s;
            regionWidth.Enabled = s;
            regionHeight.Enabled = s;
            button2.Enabled = !s;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            setCapsValue(false);
            progressBar1.Value = 1;
            oncount.RunWorkerAsync();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bgClick.CancelAsync();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (lastImg != null)
            {
                lastImg.Dispose();
            }
            if (lastb != null)
            {
                lastb.Dispose();
            }
            Bitmap b = GetScreenCapture();
            Image p = (Image)b;
            pictureBox1.BackgroundImage = p;

            lastImg = p;
            lastb = b;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m_dbConnection = new SQLiteConnection("Data Source=uiparams.db;Version=3;");
            m_dbConnection.Open();

            Dictionary<string, string> uikvs = new Dictionary<string, string>();

            string sql = "select * from param ";
            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
            SQLiteDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                uikvs[ (string)reader["keyname"] ] = (string)reader["val"];
            }

            numericUpDown1.Value = Convert.ToDecimal(uikvs["clicktimes"]);
            numericUpDown2.Value = Convert.ToDecimal(uikvs["clickperiod"]);
            numericUpDown3.Value = Convert.ToDecimal(uikvs["shotdelay"]);
            leftUpperX.Text = uikvs["upperleftx"];
            leftUpperY.Text = uikvs["upperlefty"];
            regionWidth.Text = uikvs["regionwidth"];
            regionHeight.Text = uikvs["regionheight"];
            m_dbConnection.Close();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            m_dbConnection = new SQLiteConnection("Data Source=uiparams.db;Version=3;");
            m_dbConnection.Open();
            Dictionary<string, string> uikvs = new Dictionary<string, string>();
            string sql = "select * from param ";
            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
            SQLiteDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                uikvs[(string)reader["keyname"]] = (string)reader["val"];
            }

            uikvs["clicktimes"] = numericUpDown1.Value.ToString();
            uikvs["clickperiod"] = numericUpDown2.Value.ToString();
            uikvs["shotdelay"] = numericUpDown3.Value.ToString();
            uikvs["upperleftx"] = leftUpperX.Text;
            uikvs["upperlefty"] = leftUpperY.Text;
            uikvs["regionwidth"] = regionWidth.Text;
            uikvs["regionheight"] = regionHeight.Text;
            foreach(var it in uikvs.Keys)
            {
                sql = String.Format("update param set val = '{0}' where keyname = '{1}'", uikvs[it], it);
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteReader();
            }
            m_dbConnection.Close();
        }

        private Bitmap GetScreenCapture()
        {
            Rectangle tScreenRect = new Rectangle(0, 0,Convert.ToInt32(regionWidth.Text), Convert.ToInt32(regionHeight.Text));
            Bitmap tSrcBmp = new Bitmap(tScreenRect.Width, tScreenRect.Height); 
            Graphics gp = Graphics.FromImage(tSrcBmp);
            gp.CopyFromScreen(Convert.ToInt32(leftUpperX.Text), Convert.ToInt32(leftUpperY.Text) , 0, 0, tScreenRect.Size);
            gp.DrawImage(tSrcBmp, 0, 0, tScreenRect, GraphicsUnit.Pixel);
            return tSrcBmp;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown3.Maximum = ((NumericUpDown)sender).Value;
        }
    }
}
