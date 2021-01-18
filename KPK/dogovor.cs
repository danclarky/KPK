using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPK
{
    public partial class dogovor : Form
    {
        public string photo = "";
        static readonly Random rndGen = new Random();
        public string file = "";
        public string user = "";
        public string safefile = "";
        public string audio = "";
        public dogovor()
        {
            InitializeComponent();
        }

        private void dogovor_Load(object sender, EventArgs e)
        {

            if (Environment.MachineName == "ПЛАНШЕТ")
            {
                button1.Visible = true;
            }

            this.Text = "Заявки";
            //tabControl1.TabPages[2].Parent = null;
            tabControl1.TabPages.Remove(tabPage1);

            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    if (data.typeoflist == 0)
                    {
                        cmd.CommandText = "select ФИОЗаемщика,АдресЗаемщика, ТелефонЗаемщика, РаботаЗаемщика, СуммаДоговора,ДатаДоговора,ДатаОкончания,ДатаВозврата,Договор, ДолгПоСумме, ВидОбеспечения,ФИОПоручителя,АдресПоручителя, ТелефонПоручителя,РаботаПоручителя, Подразделение,ОбъектЗалога, ДнейПросрочки,ФотоЗаемщика,Контакты,Созаемщик,ДатаРождения from Оплата where Код = '" + data.dogovortablecode + "'";
                        label6.Text = "Дата возврата";
                        dateTimePicker1.Enabled = false;
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(data.options.workmen);
                    }

                    if (data.typeoflist == 1)
                    {
                        cmd.CommandText = "select ФИОЗаемщика,АдресЗаемщика, ТелефонЗаемщика,ОстатокПоДоговору,ДатаПослПлатежа ,ДатаПослПропПлатежа,РаботаЗаемщика, СуммаДоговора,ДатаДоговора,ДатаОкончания,Договор, ДолгПоСумме, ВидОбеспечения,ФИОПоручителя,АдресПоручителя, ТелефонПоручителя,РаботаПоручителя, Подразделение,ОбъектЗалога, ДнейПросрочки,ФотоЗаемщика,Контакты,Созаемщик,ДатаРождения from Должники where Код = '" + data.dogovortablecode + "'";
                        label6.Text = "Дата платежа";
                        dateTimePicker1.Enabled = true;
                        comboBox1.Items.Clear();
                        if (data.userrules == "Юрист" || data.userrules == "Администратор")
                        {
                            comboBox1.Items.AddRange(data.options.urist);
                        }
                        else { comboBox1.Items.AddRange(data.options.worksb); }
                    }
                    if (data.typeoflist == 12)
                    {
                        cmd.CommandText = "select Договор,ФИОЗаемщика,АдресЗаемщика,ФактАдресЗаемщика, ТелефонЗаемщика,ФотоЗаемщика,ДолгКПК from УмершиеПретензии where Код = '" + data.dogovortablecode + "'";
                        Console.WriteLine(data.dogovortablecode);
                        dateTimePicker1.Enabled = true;
                        comboBox1.Items.Clear();
                        if (data.userrules == "Юрист" || data.userrules == "Администратор")
                        {
                            comboBox1.Items.AddRange(data.options.urist);
                        }
                        else { comboBox1.Items.AddRange(data.options.worksb); }
                    }
                    if (data.typeoflist == 2 || data.typeoflist == 98 || data.typeoflist == 99)
                    {
                        cmd.CommandText = "select ФИОЗаемщика,АдресЗаемщика, ТелефонЗаемщика,ДолгСуд,ДатаПлатежа ,ДатаРешения,РаботаЗаемщика, Оплатил,СуммаПлатежа,ДолгКПК,Договор, ВидОбеспечения,ФИОПоручителя,АдресПоручителя, ТелефонПоручителя,РаботаПоручителя, Подразделение,ОбъектЗалога, ФотоЗаемщика,Контакты,Созаемщик,ИП,ДатаРождения from семьшесть where Код = '" + data.dogovortablecode + "'";
                        label6.Text = "Дата платежа";
                        dateTimePicker1.Enabled = true;
                        comboBox1.Items.Clear();
                        if (data.userrules == "Юрист" || data.userrules == "Администратор")
                        {
                            comboBox1.Items.AddRange(data.options.urist);
                        }
                        else { comboBox1.Items.AddRange(data.options.worksb); }
                        tabControl1.TabPages.Add(tabPage1);
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                if (data.typeoflist == 12)
                                {
                                    textBox1.Text = reader["ФИОЗаемщика"].ToString();
                                    data.fioclient = reader["ФИОЗаемщика"].ToString();
                                    textBox2.Text = reader["Договор"].ToString();
                                    data.dogovorclient = reader["Договор"].ToString();
                                    textBox8.Text = reader["ТелефонЗаемщика"].ToString();
                                    textBox9.Text = reader["ФактАдресЗаемщика"].ToString() + " " + reader["АдресЗаемщика"].ToString();
                                    label4.Text = "Долг КПК";
                                    textBox4.Text = reader["ДолгКПК"].ToString();
                                    photo = reader["ФотоЗаемщика"].ToString();
                                }
                                else
                                {
                                    textBox1.Text = reader["ФИОЗаемщика"].ToString();
                                    textBox20.Text = reader["Датарождения"].ToString();
                                    data.fioclient = reader["ФИОЗаемщика"].ToString();
                                    textBox2.Text = reader["Договор"].ToString();
                                    data.dogovorclient = reader["Договор"].ToString();
                                    textBox8.Text = reader["ТелефонЗаемщика"].ToString() + " " + reader["Контакты"].ToString();
                                    textBox9.Text = reader["АдресЗаемщика"].ToString();
                                    textBox10.Text = reader["РаботаЗаемщика"].ToString();
                                    textBox11.Text = reader["ФИОПоручителя"].ToString();
                                    textBox12.Text = reader["ВидОбеспечения"].ToString();
                                    textBox13.Text = reader["ТелефонПоручителя"].ToString();
                                    textBox14.Text = reader["РаботаПоручителя"].ToString();
                                    textBox15.Text = reader["АдресПоручителя"].ToString();
                                    textBox16.Text = reader["ОбъектЗалога"].ToString();
                                    sozaem.Text = reader["Созаемщик"].ToString();
                                    if (data.userrules == "Менеджер" && data.typeoflist == 0)
                                    {
                                        textBox6.Text = reader["ДатаВозврата"].ToString();
                                    }
                                    else if (data.typeoflist == 1)
                                    {
                                        textBox6.Text = reader["ДатаПослПлатежа"].ToString();
                                        textBox17.Text = reader["ОстатокПоДоговору"].ToString();
                                        textBox18.Text = reader["ДатаПослПропПлатежа"].ToString();
                                    }
                                    photo = reader["ФотоЗаемщика"].ToString();
                                    this.Text = reader["ФИОЗаемщика"].ToString() + ".";
                                    if (data.typeoflist == 0)
                                    {
                                        if (data.userrules == "Менеджер" && reader["ДнейПросрочки"].ToString() == "0")
                                        {
                                            comboBox1.Items.Clear();
                                            comboBox1.Items.AddRange(data.options.workmen);
                                        }
                                    }
                                    if (data.typeoflist != 2 && data.typeoflist != 98 && data.typeoflist != 99)
                                    {
                                        textBox3.Text = reader["ДатаДоговора"].ToString();
                                        textBox4.Text = reader["СуммаДоговора"].ToString();
                                        textBox5.Text = reader["ДолгПоСумме"].ToString();
                                        textBox7.Text = reader["ДатаОкончания"].ToString();
                                    }
                                    else if (data.typeoflist == 2 || data.typeoflist == 98 || data.typeoflist == 99)
                                    {
                                        label3.Text = "Дата Решения";
                                        textBox3.Text = reader["ДатаРешения"].ToString();
                                        label4.Text = "Долг КПК";
                                        textBox4.Text = reader["ДолгКПК"].ToString();
                                        label17.Text = "Долг Суд";
                                        textBox17.Text = reader["ДолгСуд"].ToString();
                                        label6.Text = "Дата Платежа";
                                        textBox6.Text = reader["ДатаПлатежа"].ToString();
                                        label18.Text = "Сумма";
                                        textBox18.Text = reader["СуммаПлатежа"].ToString();
                                        label5.Text = "Оплатил";
                                        textBox5.Text = reader["Оплатил"].ToString();
                                    }
                                }
                            }
                        }
                        reader.Close();
                    }
                }
                conn.Close();
            }
            if (data.typeoflist == 3)
            {
                tabControl1.TabPages[0].Parent = null;
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(data.options.zayavka);
                update();
            }
            if (data.typeoflist == 11 || data.typeoflist == 13 || data.typeoflist == 14)
            {
                tabControl1.TabPages[0].Parent = null;
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(data.options.ringing);
                update();
                textBox19.Visible = true;
                label21.Visible = true;
            }
            if (data.typeoflist == 6)
            {
                textBox21.Visible = true;
                comboBox1.Visible = false;
                tabControl1.TabPages[0].Parent = null;
                update();
            }
            if (data.typeoflist == 15)
            {
                this.Text = "КонтрольКачества";
                tabControl1.Visible = false;
                panel1.Visible = true;
                this.Width = 1034;
                this.Height = 267;
                update();
            }

            String[] words = photo.Split(new char[] {';'}, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length > 0)
            {
                string qq = words[words.Length - 1];
                if (qq.Contains("10.0.0.1"))
                {
                    qq = qq.Replace("10.0.0.1", data.ipfiles);
                }
                else if (qq.Contains("192.168.1.222"))
                {
                    qq = qq.Replace("192.168.1.222", data.ipfiles);
                }
                else if (qq.Contains("78.69.157.1"))
                {
                    qq = qq.Replace("78.69.157.1", data.ipfiles);
                }
                try
                {
                    pictureBox1.Image = Image.FromFile(qq);
                }
                catch { }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string name = "";
            string ext = "";
            string path = "";
            openFileDialog1.ShowDialog();
            try
            {
                //const string rc = "йцукенгшщзхъфывапролджэячсмитьabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ0123456789";
                //for (int i = 0; i < 15; i++)
                //{
                //    name = GetRandomPassword(rc, i);
                //}
                //ext = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf(".") + 1);
                //path = "\\\\" + data.ipfiles + "\\программа\\" + name + "." + ext + "";
                //System.IO.File.Copy(openFileDialog1.FileName, path);



                Directory.CreateDirectory("\\\\" + data.ipfiles + "\\ftp\\Сканы досудебной-судебной работы\\" + ExctraxtIni(textBox1.Text));

                path = "\\\\" + data.ipfiles + "\\ftp\\Сканы досудебной-судебной работы\\" + ExctraxtIni(textBox1.Text) + "\\" + Path.GetFileName(openFileDialog1.FileName);
                System.IO.File.Copy(openFileDialog1.FileName, path, true);
                file = file + ";" + path;
                safefile = safefile + ";" + Path.GetFileName(openFileDialog1.FileName);
                String[] words = file.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                button2.Text = "Вложение     " + words.Length.ToString() + " шт.";
                MessageBox.Show("Успешно загружено");
            }
            catch { MessageBox.Show("Что то не так о_О"); }
        }

        private static string ExctraxtIni(string s)
        {
            var match = Regex.Match(s, @"(?<F>[а-яА-Я]+)(?:(?:[^а-яА-Я]+)(?<I>[а-яА-Я]+)(?:(?:[^а-яА-Я]+)(?<O>[а-яА-Я]+))?)?");
            if (!match.Success)
                return string.Empty; //подсунули дрянь :)
            var inits = match.Groups;
            if (inits["O"].Success)
                return string.Format("{0} {1}.{2}.", inits["F"], inits["I"].Value[0], inits["O"].Value[0]);
            if (inits["I"].Success)
                return string.Format("{0} {1}.", inits["F"], inits["I"].Value[0]);
            return inits["F"].Value;
        }


        static string GetRandomPassword(string ch, int pwdLength)
        {
            char[] pwd = new char[pwdLength];
            for (int i = 0; i < pwd.Length; i++)
                pwd[i] = ch[rndGen.Next(ch.Length)];
            return new string(pwd);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            data.callaudio = "";
            new attachcall().ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            data.updatetable = true;
            if ((comboBox1.SelectedIndex != -1 && textBox19.Text != "") || (textBox21.Text != "" && textBox19.Text != ""))
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;


                        if (data.typeoflist == 3)
                        {
                            cmd.CommandText = "insert into РаботаДолг (Договор,ФИО,Тип,Результат, Дата,ДатаРаботы,Путь,  Ответственный,Запись) Values( '" + textBox1.Text + "','" + textBox1.Text + "', '" + comboBox1.Text + "', '" + textBox19.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + DateTime.Now + "','" + data.dogovortablecode + "','" + data.userFIO + "','" + data.callaudio + "')";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "update Заявки set Работа = '1' where Код = '" + data.dogovortablecode + "'";
                            cmd.ExecuteNonQuery();
                        }
                        else if (data.typeoflist == 6)
                        {
                            cmd.CommandText = "insert into РаботаДолг (Договор,Тип,Результат, Дата,ДатаРаботы,  Ответственный,Запись) Values( '" + data.dogovortablecode + "','" + textBox21.Text + "', '" + textBox19.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + DateTime.Now + "','" + data.userFIO + "','" + data.callaudio + "')";
                            cmd.ExecuteNonQuery();
                        }
                        if (data.typeoflist == 11 || data.typeoflist == 13)
                        {
                            cmd.CommandText = "insert into РаботаДолг (Тип, Результат, Дата,ДатаРаботы,  Ответственный,Запись,КодДляОбзвонаКонтактов) Values( '" + comboBox1.Text + "','" + textBox19.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + DateTime.Now + "','" + data.username + "','" + data.callaudio + "','" + data.dogovortablecode + "')";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "update ОбзвонКонтакты set Работа = '" + comboBox1.Text + "' where Контакт = '" + data.dogovortablecode + "'";
                            cmd.ExecuteNonQuery();
                        }
                        if (data.typeoflist == 14)
                        {
                            cmd.CommandText = "insert into РаботаДолг (Тип, Результат, Дата,ДатаРаботы,  Ответственный,Запись,КодДляОбзвонаКонтактов) Values( '" + comboBox1.Text + "','" + textBox19.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + DateTime.Now + "','" + data.username + "','" + data.callaudio + "','" + data.dogovortablecode + "')";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "update Обзвон set Работа = '" + comboBox1.Text + "' where Контакт = '" + data.dogovortablecode + "'";
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            cmd.CommandText = "insert into РаботаДолг (Договор,ФИО,Тип,Результат, Дата,ДатаРаботы,Путь, Имя, Ответственный,Запись) Values( '" + textBox2.Text + "','" + textBox1.Text + "', '" + comboBox1.Text + "', '" + textBox19.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + DateTime.Now + "','" + file + "','" + safefile + "','" + data.userFIO + "','" + data.callaudio + "')";
                            cmd.ExecuteNonQuery();
                            if (data.typeoflist == 0)
                            {
                                cmd.CommandText = "update Оплата set Пропуск = '" + dateTimePicker1.Value.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                            }
                            if (data.typeoflist == 1)
                            {
                                cmd.CommandText = "update Должники set Пропуск = '" + dateTimePicker1.Value.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                            }
                            if (data.typeoflist == 2 || data.typeoflist == 98 || data.typeoflist == 99)
                            {
                                cmd.CommandText = "update семьшесть set Пропуск = '" + dateTimePicker1.Value.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                            }
                            if (data.typeoflist == 12)
                            {
                                cmd.CommandText = "update УмершиеПретензии set Пропуск = '" + dateTimePicker1.Value.ToShortDateString() + "' where Код = '" + data.dogovortablecode + "'";
                            }
                            cmd.ExecuteNonQuery();
                        }
                    }
                    conn.Close();
                }
                file = "";
                safefile = "";
                audio = "";
                comboBox1.SelectedIndex = -1;
                textBox21.Text = "";
                textBox19.Text = "";
                button2.Text = "Вложение";
                update();
            }
            else
            {
                MessageBox.Show("Выберите тип работы и введите результат");
            }
            data.callaudio = "";
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                update();
            }
            if (tabControl1.SelectedIndex == 2)
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        dataGridView2.Rows.Clear();
                        cmd.CommandText = "select Дата,Основной,Проценты, Членские,Штрафы from ИсторияРасчетов where Договор = '" + textBox2.Text + "' and ФИО = '" + textBox1.Text + "' and (Основной<>'0' or Проценты<>'0' or Членские<>'0' or Штрафы<>'0') order by Дата";
                        using (var reader = cmd.ExecuteReader())
                        {
                            int i = 0;
                            while (reader.Read())
                            {

                                dataGridView2.Rows.Add();
                                dataGridView2.Rows[i].Cells[0].Value = reader["Дата"].ToString();
                                dataGridView2.Rows[i].Cells[1].Value = reader["Основной"].ToString();
                                dataGridView2.Rows[i].Cells[2].Value = reader["Проценты"].ToString();
                                dataGridView2.Rows[i].Cells[3].Value = reader["Членские"].ToString();
                                dataGridView2.Rows[i].Cells[4].Value = reader["Штрафы"].ToString();

                                i++;
                            }
                            reader.Close();
                        }
                    }
                }
            }
            if (tabControl1.SelectedIndex == 3)
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        dataGridView3.Rows.Clear();
                        cmd.CommandText = "select ИПШКИ from ИПСУД where Владелец = '" + textBox1.Text + "'";
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            int i = 0;
                            string text = "";
                            while (reader.Read())
                            {

                                text = reader["ИПШКИ"].ToString();
                                string[] IPText = text.Split('~');
                                foreach (string element in IPText)
                                {
                                    dataGridView3.Rows.Add();
                                    string[] IPElement = element.Split(';');
                                    dataGridView3.Rows[i].Cells[0].Value = IPElement[0];
                                    if (IPElement[2] == "Нет")
                                    { 
                                        dataGridView3.Rows[i].Cells[1].Value = IPElement[2]; 
                                    }
                                    else
                                    {
                                        dataGridView3.Rows[i].Cells[1].Value = IPElement[2].Split(',')[0];
                                    }
                                    dataGridView3.Rows[i].Cells[2].Value = IPElement[4];
                                    dataGridView3.Rows[i].Cells[3].Value = IPElement[3];
                                    i++;
                                }
                                
                            }
                            reader.Close();
                        }
                    }
                }
            }

        }

        void update()
        {

            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    dataGridView1.Rows.Clear();
                    if (data.typeoflist != 15)
                    {
                        if (data.typeoflist == 3)
                        {
                            cmd.CommandText = "select Код,Тип,Результат, Дата,Ответственный,Путь,Имя,Запись from РаботаДолг where Путь = '" + data.dogovortablecode + "' order by ДатаРаботы";

                        }
                        else if (data.typeoflist == 11 || data.typeoflist == 13 || data.typeoflist == 14)
                        {
                            DateTime date1 = new DateTime(2020, 05, 01);
                            cmd.CommandText = "select Код,Тип,Результат, Дата,Ответственный,Путь,Имя,Запись from РаботаДолг where КодДляОбзвонаКонтактов = '" + data.dogovortablecode + "' and ДатаРаботы > '" + date1 + "' order by ДатаРаботы";
                       
                        }
                        else if (data.typeoflist == 6)
                        {
                            cmd.CommandText = "select Код,Тип,Результат, Дата,Ответственный,Путь,Имя,Запись from РаботаДолг where Договор = '" + data.dogovortablecode + "' order by ДатаРаботы";
                        }
                        else
                        {
                            cmd.CommandText = "select Код,Тип,Результат, Дата,Путь, Имя, Ответственный,Запись from РаботаДолг where Договор = '" + textBox2.Text + "' and ФИО ='" + textBox1.Text + "' order by ДатаРаботы";
                        }
                        using (var reader = cmd.ExecuteReader())
                        {
                            int i = 0;
                            while (reader.Read())
                            {
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows[i].Cells[0].Value = reader["Код"].ToString();
                                dataGridView1.Rows[i].Cells[1].Value = reader["Тип"].ToString();
                                dataGridView1.Rows[i].Cells[2].Value = reader["Дата"].ToString();
                                dataGridView1.Rows[i].Cells[3].Value = reader["Результат"].ToString();
                                dataGridView1.Rows[i].Cells[4].Value = reader["Ответственный"].ToString();
                                if (reader["Имя"].ToString() != "") dataGridView1.Rows[i].Cells[5].Value = "Вложение";
                                else dataGridView1.Rows[i].Cells[5].Value = "X";
                                if (reader["Запись"].ToString() != "") dataGridView1.Rows[i].Cells[6].Value = "Аудио";
                                else dataGridView1.Rows[i].Cells[6].Value = "X";
                                i++;
                            }
                            reader.Close();
                        }
                    }
                    else if (data.typeoflist == 15)
                    {
                        cmd.CommandText = "SELECT Контрагент,Дата,Контакт,Сумма,Вопрос1,Вопрос2,Вопрос3 FROM КонтрольКачестваЗвонки where Код = '" + data.dogovortablecode + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                textBox29.Text = reader["Сумма"].ToString();
                                textBox22.Text = reader["Дата"].ToString();
                                textBox28.Text = reader["Контрагент"].ToString();
                                textBox24.Text = reader["Контакт"].ToString();
                                comboBox3.Text = reader["Вопрос1"].ToString();
                                comboBox2.Text = reader["Вопрос2"].ToString();
                                comboBox4.Text = reader["Вопрос3"].ToString();
                            }
                            reader.Close();
                        }
                    }
                }
            }
        }



        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select Результат from РаботаДолг where Код = '" + dataGridView1.Rows[e.RowIndex].Cells[0].Value + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                MessageBox.Show(reader["Результат"].ToString());
                            }
                            reader.Close();
                        }
                    }
                }
            }
            if (e.ColumnIndex == 5)
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select Путь,Имя from РаботаДолг where Код = '" + dataGridView1.Rows[e.RowIndex].Cells[0].Value + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                data.attachfile = reader["Путь"].ToString();
                                data.namefile = reader["Имя"].ToString();
                            }
                            reader.Close();
                        }
                    }
                }
                if (data.namefile != "")
                {
                    new attach().ShowDialog();
                }
            }
            if (e.ColumnIndex == 6)
            {
                string audio = "";
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select Запись from РаботаДолг where Код = '" + dataGridView1.Rows[e.RowIndex].Cells[0].Value + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                audio = reader["Запись"].ToString();
                            }
                            reader.Close();
                        }
                    }
                }
                if (audio != "")
                {
                  
                else { MessageBox.Show("нету аудио"); }
            }
        }

        private void dogovor_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            String[] words = photo.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length > 0)
            {
                string qq = words[words.Length - 1];
        
                System.Diagnostics.Process.Start(qq);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new foto().ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex != -1 && comboBox4.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)
            {
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "update КонтрольКачестваЗвонки set Вопрос1 = '" + comboBox3.Text + "',Вопрос2 = '" + comboBox2.Text + "',Вопрос3 = '" + comboBox4.Text + "', Аудио='" + data.callaudio + "', Ответственный = '" + data.userFIO + "' where Код = '" + data.dogovortablecode + "'";
                        cmd.ExecuteNonQuery();

                    }
                }
                data.updatetable = true;
            }
            else 
            {
                MessageBox.Show("Укажите все пункты");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            data.callaudio = "";
            new attachcall().ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string audio = "";
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select Аудио from КонтрольКачестваЗвонки where Код = '" + data.dogovortablecode + "'";
                    Console.WriteLine(cmd.CommandText);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            audio = reader["Аудио"].ToString();
                        }
                        reader.Close();
                    }
                }
            }
            if (audio != "")
            {
                
                System.Diagnostics.Process.Start(audio);
            }
        }

    }
}
