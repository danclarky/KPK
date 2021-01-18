using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace KPK
{
    public partial class notification : Form
    {
        public notification()
        {
            InitializeComponent();
            this.Height = Convert.ToInt32(SystemParameters.WorkArea.Height) / 2;
            this.Location = new System.Drawing.Point(Convert.ToInt32(SystemParameters.WorkArea.Width) - this.Width, Convert.ToInt32(SystemParameters.WorkArea.Height) - this.Height);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            timer1.Start();
            webBrowser1.DocumentText = "<html><style> .head{background:#33FFCC;text-align:center;}.txt{text-align:center;color:white;} body{background:red} label{text-align:center;} p{border-style:double;margin:0;}</style><body><br>" + data.options.notiftext + "</body></html>";
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
