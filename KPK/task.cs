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
    public partial class task : Form
    {
        string owner = null,ispolnitel = null;
        int choose = 0;
        public task()
        {
            InitializeComponent();
        }

        private void textBox2_MouseClick(object sender, MouseEventArgs e)
        {
            dataGridView1.Visible = false; 
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String[] words = textBox2.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string query = "";
            dataGridView1.Visible = true;
            dataGridView1.Rows.Clear();
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    query = "select Код, ФИО from Пользователи where Логин <> 'Администратор' ";
                    for (int i = 0; i != words.Length; i++)
                    {
                        query += "and ФИО <> '" + words [i]+ "'";
                    }
                    query += "order by ФИО";
                    cmd.CommandText = query;
                    using (var reader = cmd.ExecuteReader())
                    {
                        int i = 0;
                        while (reader.Read())
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[i].Cells[0].Value = reader["ФИО"].ToString();
                            dataGridView1.Rows[i].Cells[1].Value = reader["Код"].ToString();
                            i++;
                        }
                        reader.Close();
                    }
                }
            }
            toolStripMenuItem1.Text = "Добавить";
            choose = 1;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Visible = true;
            String[] words = textBox2.Text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i != words.Length; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = words[i];
            }
            toolStripMenuItem1.Text = "Удалить";
            choose = 2;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string txt = "";
            if (choose == 1)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1[0, i].Selected)
                        textBox2.Text += dataGridView1[0, i].Value.ToString() + ";";
                }
            }
            else if (choose == 2)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1[0, i].Selected)
                        txt += dataGridView1[0, i].Value.ToString() + ";";
                }
                textBox2.Text = txt;
            }
            dataGridView1.Visible = false;
        }

        private void task_Load(object sender, EventArgs e)
        {
            if (data.typeactiontask == "newtask")
            {
                checkBox1.Visible = false;
                textBox1.Text = data.userFIO;
            }
            if (data.typeactiontask == "opentask")
            {
                checkBox1.Visible = true;
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select Владелец,Исполнитель from Задания where Код = '"+data.idtask+"' ";
                        using (var reader = cmd.ExecuteReader())
                        {
                            int i = 0;
                            while (reader.Read())
                            {
                                owner = reader["Владелец"].ToString();
                                ispolnitel = reader["Исполнитель"].ToString();
                                i++;
                            }
                            reader.Close();
                        }

                        if (owner == data.userFIO )
                        {
                            checkBox1.Visible = false;
                        }

                    }
                }
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
                    if (data.options.typeoflist != 5)
                    {
                        cmd.CommandText = "insert into Задания (Тема,Владелец,Исполнитель,Описание,ДатаНачала,ДатаОкончания) Values( '" + textBox4.Text + "','" + textBox1.Text + "', '" + textBox2.Text + "', '" + textBox3.Text + "','"+DateTime.Now+"','" + dateTimePicker1.Value + "','0','0')";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

    }
}
