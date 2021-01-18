using AsterNET.Manager;
using AsterNET.Manager.Action;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPK
{
    public partial class adding : Form
    {
        public adding()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (data.adiingtype == "cons")
            {
                label3.Text = "Телефон";
            }
            else if (data.adiingtype == "sber")
            {
                label2.Visible = true;
                dateTimePicker1.Visible = true;
                label3.Text = "Сумма выдачи";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    if (data.adiingtype == "cons")
                    {
                        cmd.CommandText = "insert into Консультации (ФИО,Телефон,Подразделение,Менеджер,Дата) Values( '" + textBox1.Text + "','" + textBox3.Text + "', '" +data.usercity + "', '" + data.userFIO + "','" + DateTime.Now + "')";
                        cmd.ExecuteNonQuery();
                    }
                    else if (data.adiingtype == "sber")
                    {
                        cmd.CommandText = "insert into Сбережения (ФИОЗаемщика,СуммаВыплаты,Подразделение,ДатаВыплаты,Выдан) Values( '" + textBox1.Text + "','" + textBox3.Text + "', '" + data.usercity + "', '" + dateTimePicker1.Value + "','Нет')";
                        cmd.ExecuteNonQuery();
                    }
                   
                }
                conn.Close();
            }
            this.Close();
        }

        private void adding_FormClosing(object sender, FormClosingEventArgs e)
        {
            data.updatetable = true;
        }
      
    }
}
