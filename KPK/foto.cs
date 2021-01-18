using AForge.Video.DirectShow;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPK
{
    public partial class foto : Form
    {
        static readonly Random rndGen = new Random();
        public string file = "";
        public string safefile = "";
        public foto()
        {
            InitializeComponent();
        }
        private FilterInfoCollection CaptureDevice;
        private VideoCaptureDevice FinalFrame;

        private void foto_Load(object sender, EventArgs e)
        {
            CaptureDevice = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            FinalFrame = new VideoCaptureDevice();
            foreach (FilterInfo Device in CaptureDevice)
            {
                comboBox1.Items.Add(Device.Name);
            }
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        private void foto_FormClosed(object sender, FormClosedEventArgs e)
        {
            videoSourcePlayer3.Stop();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            videoSourcePlayer3.Stop();
            videoSourcePlayer3.VideoSource = new VideoCaptureDevice(CaptureDevice[comboBox1.SelectedIndex].MonikerString);
            videoSourcePlayer3.Start();
        }
        static string GetRandomPassword(string ch, int pwdLength)
        {
            char[] pwd = new char[pwdLength];
            for (int i = 0; i < pwd.Length; i++)
                pwd[i] = ch[rndGen.Next(ch.Length)];
            return new string(pwd);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string name = "";
            string path = "";


            const string rc = "йцукенгшщзхъфывапролджэячсмитьabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ0123456789";
            for (int i = 0; i < 15; i++)
            {
                name = GetRandomPassword(rc, i);
            }

            path = "\\\\" + data.ipfiles + "\\программа\\" + name + ".jpg";
            file = file + ";" + path;
            safefile = safefile + ";" + name;
           
            Bitmap varBmp = videoSourcePlayer3.GetCurrentVideoFrame();
            varBmp.Save(path, ImageFormat.Jpeg);
            varBmp.Dispose();
            varBmp = null;
            



        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;



                    cmd.CommandText = "insert into РаботаДолг (Договор,ФИО,Тип, Дата,ДатаРаботы,Путь, Имя, Ответственный) Values( '" + data.dogovorclient + "','" + data.fioclient + "', '" + comboBox2.Text + "','" + DateTime.Now.ToShortDateString() + "','" + DateTime.Now + "','" + file + "','" + safefile + "','" + data.userFIO + "')";
                    cmd.ExecuteNonQuery();
                    if (data.typeoflist == 0)
                    {
                        cmd.CommandText = "update Оплата set Пропуск = '" + DateTime.Now.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                    }
                    if (data.typeoflist == 1)
                    {
                        cmd.CommandText = "update Должники set Пропуск = '" + DateTime.Now.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                    }
                    if (data.typeoflist == 2)
                    {
                        cmd.CommandText = "update семьшесть set Пропуск = '" + DateTime.Now.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                    }
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
            this.Close();
        }
    }
}
