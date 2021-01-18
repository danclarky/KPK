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
    public partial class select : Form
    {
        public select()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void select_Load(object sender, EventArgs e)
        {

            comboBox4.Items.Clear();
           comboBox4.Items.AddRange(data.options.users());
            if (data.checkaction == "edit")
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select ТипПроверки,ФИОСотрудника,Проверяющий,Проблема,ОтветПроверяющего,ОтветМенеджера,ДатаОтветаПроверяющего,ДатаОтветаМенеджера from ПроверкаМенеджеров where Код = '" + data.checkcode + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                comboBox4.Text = reader["ФИОСотрудника"].ToString();
                                comboBox5.Text = reader["ТипПроверки"].ToString();
                                textBox1.Text = reader["Проблема"].ToString();
                                try
                                {
                                    dateTimePicker1.Value = Convert.ToDateTime(reader["ДатаОтветаПроверяющего"]);
                                    dateTimePicker2.Value = Convert.ToDateTime(reader["ДатаОтветаМенеджера"]);
                                }
                                catch { }

                                if (data.userrules == "Менеджер")
                                {
                                    comboBox4.Enabled = false;
                                    button4.Visible = false;
                                    textBox1.Enabled = false;
                                    comboBox5.Enabled = false;
                                    dateTimePicker1.Visible = false;
                                    dateTimePicker2.Visible = false;
                                }
                                if (reader["ОтветПроверяющего"].ToString() == "Да")
                                {
                                    button7.Text = "Убрать";

                                }
                                else
                                {
                                    button7.Text = "Проверить";
                                }
                                if (reader["ОтветМенеджера"].ToString() == "Да")
                                {
                                    button4.Text = "Убрать за менеджера";
                                    if (data.userrules == "Менеджер")
                                    {
                                        button7.Visible = false;
                                    }
                                }
                                else
                                {
                                    button4.Text = "Проверить за менеджера";
                                }
                            }
                            reader.Close();
                        }
                    }

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton2.Checked && data.options.userrules == "Юрист")
            {
                data.options.typequery = 1;
                data.options.query = "select Код, ФИОЗаемщика, Договор,Подразделение, СуммаДоговора,Пропуск, ДатаДоговора,ДатаОкончания,ДатаПослПлатежа,ДатаПослПропПлатежа,ДолгПоСумме,Просрочка,ОстатокПоДоговору from Должники where Просрочка > '59' and Подразделение = '" + comboBox1.Text + "' order by Просрочка desc";
            }
            if (radioButton2.Checked && (data.options.userrules == "Админ" || data.options.userrules == "НСБ"))
            {
                data.options.typequery = 1;
                data.options.query = " where Подразделение = '" + comboBox1.Text + "' ";
            }
            if (radioButton3.Checked && (data.options.userrules == "Админ" || data.options.userrules == "НСБ"))
            {
                data.options.typequery = 1;
                data.options.query = " where Специалист = '" + comboBox3.Text + "' ";
            }
            //if (radioButton2.Checked && (data.options.userrules == "Админ" || data.options.userrules == "НСБ") && data.options.typeoflist == 1)
            //{
            //    data.options.typequery = 1;
            //    data.options.query = " where Подразделение = '" + comboBox1.Text + "' ";
            //}
            if (radioButton2.Checked)
            {
                data.options.typesort = "Подразделение";
                data.options.textsort = comboBox1.Text;
            }

            if (radioButton3.Checked)
            {

                data.options.typesort = "Сотрудник";
                data.options.textsort = comboBox3.Text;

            }

            this.Close();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked) comboBox1.Enabled = true;
            else { comboBox1.Enabled = false; comboBox3.Enabled = false; }
        }



        private void button3_Click(object sender, EventArgs e)
        {
            data.options.typequery = 0;
            data.options.query = "";
            data.options.typesort = "";
            data.options.textsort = "";
            this.Close();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked) comboBox3.Enabled = true;
            else { comboBox1.Enabled = false; }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (data.userrules == "Менеджер")
            { this.Close(); }
            else
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;

                        cmd.CommandText = "select Подразделение from Пользователи where ФИО = '" + comboBox4.Text + "'";
                        string fiomen = "";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                fiomen = reader["Подразделение"].ToString();
                            }
                            reader.Close();
                        }
                        if (data.checkaction == "edit")
                        {
                            cmd.CommandText = "update ПроверкаМенеджеров set Подразделение = '" + fiomen + "',ТипПроверки = '" + comboBox5.Text + "',ФИОСотрудника = '" + comboBox4.Text + "',Проблема = '" + textBox1.Text + "' where Код = '" + data.checkcode + "'";
                        }
                        if (data.checkaction == "add")
                        {
                            cmd.CommandText = "insert into ПроверкаМенеджеров (Подразделение,ТипПроверки,Дата,ФИОСотрудника,Проверяющий,Проблема,ОтветМенеджера,ОтветПроверяющего) Values( '" + fiomen + "','" + comboBox5.Text + "', '" + DateTime.Now + "', '" + comboBox4.Text + "','" + data.userFIO + "','" + textBox1.Text + "','Нет','Нет')";
                        }
                        cmd.ExecuteNonQuery();


                    }
                    conn.Close();
                }
                this.Close();
            }
            data.checkcode = "";
            data.updatetable = true;
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    if (button7.Text == "Проверить")
                    {
                        if (data.userrules == "Менеджер")
                        {
                            cmd.CommandText = "update ПроверкаМенеджеров set ОтветМенеджера = 'Да', ДатаОтветаМенеджера = '" + dateTimePicker1.Value + "' where Код = '" + data.checkcode + "'"; 
                            data.updatetable = true;
                            this.Close();
                        }
                        else
                        {
                            cmd.CommandText = "update ПроверкаМенеджеров set ОтветПроверяющего = 'Да', ДатаОтветаПроверяющего = '" + dateTimePicker1.Value + "' where Код = '" + data.checkcode + "'";
                        }
                        button7.Text = "Убрать";
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    else
                    {
                        cmd.CommandText = "update ПроверкаМенеджеров set ОтветПроверяющего = 'Нет' where Код = '" + data.checkcode + "'";
                        button7.Text = "Проверить";
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }


                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    if (button4.Text == "Проверить за менеджера")
                    {
                        cmd.CommandText = "update ПроверкаМенеджеров set ОтветМенеджера = 'Да', ДатаОтветаМенеджера = '" + dateTimePicker2.Value + "' where Код = '" + data.checkcode + "'";
                        button4.Text = "Убрать за менеджера";
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    else
                    {
                        cmd.CommandText = "update ПроверкаМенеджеров set ОтветМенеджера = 'Нет' where Код = '" + data.checkcode + "'";
                        button4.Text = "Проверить за менеджера";
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                }
            }
        }
    }
}
