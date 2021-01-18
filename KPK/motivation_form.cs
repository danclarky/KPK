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
    public partial class motivation_form : Form
    {
        public motivation_form()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1)
            {

                foreach (TextBox c in tabPage1.Controls.OfType<TextBox>())
                {
                    if (c.Text == "")
                    {
                        c.Text = "0";
                    }
                }
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "Insert into Мотивация_План (Подразделение,Дата,Быстрый,БыстрыйПлас,Потреб,Иное,МСК,Пайщики,Оплата,Закрытие,Претензии,Волна) values('" + comboBox1.Text + "','" + dateTimePicker1.Value.Date + "','" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + textBox9.Text + "','" + textBox10.Text + "')";
                        cmd.ExecuteNonQuery();
                    }
                }
                update();
                foreach (TextBox c in tabPage1.Controls.OfType<TextBox>())
                {
                    c.Text = "";
                }

            }
            else
            {
                MessageBox.Show("Выберите Подразделение");
            }
        }


        public void update()
        {

            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "SELECT Код,Подразделение,Дата,Быстрый,БыстрыйПлас,Потреб,МСК,Иное,Пайщики,Оплата,Закрытие,Претензии,Волна FROM Мотивация_План";
                    motiv_plan.Rows.Clear();
                    using (var reader = cmd.ExecuteReader())
                    {
                        int i = 0;
                        while (reader.Read())
                        {
                            motiv_plan.Rows.Add();
                            motiv_plan.Rows[i].Cells[0].Value = i + 1;
                            motiv_plan.Rows[i].Cells[1].Value = reader["Код"].ToString();
                            motiv_plan.Rows[i].Cells[2].Value = reader["Дата"].ToString();
                            motiv_plan.Rows[i].Cells[3].Value = reader["Подразделение"].ToString();
                            motiv_plan.Rows[i].Cells[4].Value = reader["Быстрый"].ToString();
                            motiv_plan.Rows[i].Cells[5].Value = reader["БыстрыйПлас"].ToString();
                            motiv_plan.Rows[i].Cells[6].Value = reader["Потреб"].ToString();
                            motiv_plan.Rows[i].Cells[7].Value = reader["Иное"].ToString();
                            motiv_plan.Rows[i].Cells[8].Value = reader["МСК"].ToString();
                            motiv_plan.Rows[i].Cells[9].Value = reader["Пайщики"].ToString();
                            motiv_plan.Rows[i].Cells[10].Value = reader["Оплата"].ToString();
                            motiv_plan.Rows[i].Cells[11].Value = reader["Закрытие"].ToString();
                            motiv_plan.Rows[i].Cells[12].Value = reader["Претензии"].ToString();
                            motiv_plan.Rows[i].Cells[13].Value = reader["Волна"].ToString();
                            i++;
                        }
                        reader.Close();
                    }
                    motiv_control.Rows.Clear();
                    cmd.Connection = conn;
                    cmd.CommandText = "SELECT Код,Дата,Сотрудник,Тип,Значение FROM Мотивация_Контроль";
                    using (var reader = cmd.ExecuteReader())
                    {
                        int i = 0;
                        while (reader.Read())
                        {
                            motiv_control.Rows.Add();
                            motiv_control.Rows[i].Cells[0].Value = i + 1;
                            motiv_control.Rows[i].Cells[1].Value = reader["Код"].ToString();
                            motiv_control.Rows[i].Cells[2].Value = reader["Дата"].ToString();
                            motiv_control.Rows[i].Cells[3].Value = reader["Сотрудник"].ToString();
                            motiv_control.Rows[i].Cells[4].Value = reader["Тип"].ToString();
                            motiv_control.Rows[i].Cells[5].Value = reader["Значение"].ToString();
                            i++;
                        }
                        reader.Close();
                    }
                }
            }
        }



        private void motivation_form_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(data.options.DOMotiv());
            comboBox2.Items.AddRange(data.options.UsersMotiv());
            update();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1 || comboBox3.SelectedIndex != -1)
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "Insert into Мотивация_Контроль (Сотрудник,Дата,Тип,Значение) values('" + comboBox2.Text + "','" + dateTimePicker2.Value.Date + "','" + comboBox3.Text + "','" + textBox11.Text + "')";
                        cmd.ExecuteNonQuery();
                    }
                }
                update();
                foreach (TextBox c in tabPage2.Controls.OfType<TextBox>())
                {
                    c.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Выберите сотрудника или тип");
            }
        }
    }
}
