using MySql.Data.MySqlClient;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPK
{
    public partial class analiz : Form
    {
        public analiz()
        {
            InitializeComponent();
        }

        private void users_Load(object sender, EventArgs e)
        {
            if (data.whatanaliz == "проверка")
            {
                this.Text = "Аналитика последконтроль";
                dateend.Value = DateTime.Now;
                datestart.Value = DateTime.Now.AddDays(-7);
                tableofcheckanaliz.Visible = true;
                tableofcallsanaliz.Visible = false;
            }
            if (data.whatanaliz == "телефония")
            {
                dateend.Visible = false;
                this.Text = "Аналитика телефония";
                datestart.Value = DateTime.Now;
                tableofcallsanaliz.Visible = true;
                tableofcheckanaliz.Visible = false;
            }
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (data.whatanaliz == "телефония")
            {

                tableofcallsanaliz.Rows.Clear();
                string[,] DOO;
                int[] allisx;
                int[] allvxod;
                int[] otvisx;
                int[] otvvxod;
                int[] nototvisx;
                int[] nototvvxod;
                int[] timevxod;
                int[] timeisxod;
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select count(Код) from ДОms";
                        object value = cmd.ExecuteScalar();
                        DOO = new string[Convert.ToInt32(value), 2];
                        allisx = new int[Convert.ToInt32(value)];
                        allvxod = new int[Convert.ToInt32(value)];
                        otvisx = new int[Convert.ToInt32(value)];
                        otvvxod = new int[Convert.ToInt32(value)];
                        nototvisx = new int[Convert.ToInt32(value)];
                        nototvvxod = new int[Convert.ToInt32(value)];
                        timevxod = new int[Convert.ToInt32(value)];
                        timeisxod = new int[Convert.ToInt32(value)];
                        cmd.CommandText = "select ДО,Номер from ДОms";
                        int i = 0;
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                DOO[i, 0] = reader["ДО"].ToString();
                                DOO[i, 1] = reader["Номер"].ToString();

                                i++;

                            }
                        }
                    }
                }
                string datemysqlfield = datestart.Value.ToString().Substring(6, 4) + "-" + datestart.Value.ToString().Substring(3, 2) + "-" + datestart.Value.ToString().Substring(0, 2);
                using (var conn = new MySqlConnection(data.stringconnect()[0]))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand())
                    {
                        cmd.Connection = conn;
                        for (int q = 0; q < DOO.GetLength(0); q++)
                        {
                            cmd.CommandText = "SELECT src,dst ,disposition ,calldate,billsec,recordingfile,uniqueid FROM cdr where LEFT(CAST(calldate as char), 10) = '" + datemysqlfield + "'";
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    if (reader["src"].ToString() == DOO[q, 1])
                                    {
                                        timeisxod[q] += Convert.ToInt32(reader["billsec"]);
                                        allisx[q]++;
                                        if (reader["disposition"].ToString() == "ANSWERED") otvisx[q]++;
                                        else nototvisx[q]++;


                                    }
                                    if (reader["dst"].ToString() == DOO[q, 1])
                                    {
                                        timevxod[q] += Convert.ToInt32(reader["billsec"]);
                                        allvxod[q]++;
                                        if (reader["disposition"].ToString() == "ANSWERED") otvvxod[q]++;
                                        else nototvvxod[q]++;


                                    }

                                }
                            }
                            tableofcallsanaliz.Rows.Add();
                            tableofcallsanaliz.Rows[q].Cells[0].Value = DOO[q, 0];
                            tableofcallsanaliz.Rows[q].Cells[1].Value = allvxod[q] + "|" + otvvxod[q] + "|" + nototvvxod[q] + "|" + timevxod[q]+"c.";
                            tableofcallsanaliz.Rows[q].Cells[2].Value = allisx[q] + "|" + otvisx[q] + "|" + nototvisx[q] + "|" + timeisxod[q]+"c.";
                        }
                    }
                }
            }

            if (data.whatanaliz == "проверка")
            {

                tableofcheckanaliz.Rows.Clear();
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select count(Код) from Пользователи where Долж = 'Менеджер'";
                        object value = cmd.ExecuteScalar();
                        string[] userfio = new string[Convert.ToInt32(value)];
                        int[] userallmistakes = new int[Convert.ToInt32(value)];
                        int[] userisprmistakes = new int[Convert.ToInt32(value)];
                        int[] userignoredmistakes = new int[Convert.ToInt32(value)];
                        int[] usercostsmistakes = new int[Convert.ToInt32(value)];
                        cmd.CommandText = "select ФИО from Пользователи where Долж = 'Менеджер' order by ФИО ";
                        int i = 0;
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                userfio[i] = reader["ФИО"].ToString();
                                i++;
                            }
                        }
                        for (int q = 0; q < userfio.Length; q++)
                        {
                            cmd.CommandText = "SELECT Дата,ФИОСотрудника,ОтветПроверяющего,ДатаОтветаПроверяющего FROM ПроверкаМенеджеров where ФИОСотрудника = '" + userfio[q] + "' and (Дата < '" + dateend.Value.AddDays(1) + "' and Дата >= '" + datestart.Value.AddDays(-1) + "')";
                            int relative1 = 0;
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    i = 0;
                                    while (reader.Read())
                                    {
                                        if (reader["ОтветПроверяющего"].ToString() == "Да") userisprmistakes[q]++;
                                        else userignoredmistakes[q]++;
                                        if (reader["ДатаОтветаПроверяющего"] != DBNull.Value)
                                        {
                                            TimeSpan span1 = Convert.ToDateTime(reader["ДатаОтветаПроверяющего"]) - Convert.ToDateTime(reader["Дата"]);
                                            relative1 = span1.Days;
                                        }
                                        else
                                        {
                                            TimeSpan span1 = DateTime.Now - Convert.ToDateTime(reader["Дата"]);
                                            relative1 = span1.Days;
                                        }
                                        if (relative1 > 0) usercostsmistakes[q] += 50;
                                        usercostsmistakes[q] += relative1 * 20;
                                        i++;
                                    }
                                    userallmistakes[q] = i;
                                }
                            }
                            tableofcheckanaliz.Rows.Add();
                            tableofcheckanaliz.Rows[q].Cells[0].Value = userfio[q];
                            tableofcheckanaliz.Rows[q].Cells[1].Value = userallmistakes[q] + "|" + userisprmistakes[q] + "|" + userignoredmistakes[q];
                            tableofcheckanaliz.Rows[q].Cells[2].Value = usercostsmistakes[q] + "p.";
                        }

                    }
                }


            }

        }
    }
}
