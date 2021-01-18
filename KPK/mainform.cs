using AsterNET.Manager;
using AsterNET.Manager.Action;
using MySql.Data.MySqlClient;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Text.RegularExpressions;
using AsterNET.Manager.Event;



namespace KPK
{
    public partial class mainform : Form
    {
        int sum = 0;
        public string whatfind = "", limit = "", refreshpages = "0", fio = "", podr = "";
        ToolStripButton findcancel = new ToolStripButton();
        ToolStripButton findok = new ToolStripButton();
        ToolStripComboBox findfootpodr = new ToolStripComboBox();
        ToolStripStatusLabel label2 = new ToolStripStatusLabel();
        ToolStripTextBox findfoottxt = new ToolStripTextBox();
        ToolStripStatusLabel label1 = new ToolStripStatusLabel();
        ToolStripStatusLabel label4 = new ToolStripStatusLabel();
        ToolStripStatusLabel labelfoot = new ToolStripStatusLabel();
        private string CodeToAuthenticate { get; set; }
        public Double alldolg = 0, obzvonmanager = 0;
        public int napravcalls;
        public string[] selectcalls = new string[3];
        public string[] selectcheck = new string[3];
        public string[] selectoldbase = new string[3];
        static string datemysql = DateTime.Now.ToShortDateString().Substring(6, 4) + "-" + DateTime.Now.ToShortDateString().Substring(3, 2) + "-" + DateTime.Now.ToShortDateString().Substring(0, 2);
        public string selectofcheck = "", selectofoldbase = "", datecall = "and LEFT(CAST(calldate as char), 10) = '" + datemysql + "'", selectcall = "";
        public DataGridView table;



        public mainform()
        {
            SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;
            InitializeComponent();
        }

        void oplata(string dogovorcode, NpgsqlCommand cmd)
        {
            string txt = "select Код, ФИОЗаемщика,Подразделение,Пропуск, Договор, СуммаДоговора,Пропуск, ДатаВозврата,ДатаДоговора,ДатаОкончания,ДолгПоСумме,ДнейПросрочки, (select Тип from РаботаДолг where ФИО = ФИОЗаемщика and РаботаДолг.Договор=Оплата.Договор order by ДатаРаботы desc Limit 1) as notif from Оплата";

            if ((data.userrules == "Администратор" || data.userrules == "НСБ" || data.userrules == "ГлавМенеджер"))
            {
                if (DateTime.Today == new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1))
                {
                    txt += " where (ДатаВозврата <= '" + DateTime.Today + "' or ДнейПросрочки <> '0') " + data.selectoplata + " order by ДатаВозврата";
                }
                else
                {
                    txt += " where (ДатаВозврата <= '" + DateTime.Today.AddDays(1) + "' or ДнейПросрочки <> '0') " + data.selectoplata + " order by ДатаВозврата";
                }
            }
            else
            {
                String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                txt += " where (Подразделение = '" + data.usercity + "' ";
                for (int i = 0; i != words.Length; i++)
                {
                    txt += " or Подразделение = '" + words[i] + "'";
                }
                txt += " ) order by ДнейПросрочки desc,ДатаВозврата";
            }

            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }
            visible(listofoplata, q);
            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    string dolg = "";
                    listofoplata.Rows.Add();
                    listofoplata.Rows[i].Cells[0].Value = i + 1;
                    listofoplata.Rows[i].Cells[1].Value = reader["Код"].ToString();
                    listofoplata.Rows[i].Cells[2].Value = reader["Договор"].ToString();
                    listofoplata.Rows[i].Cells[3].Value = reader["ФИОЗаемщика"].ToString();
                    if (reader["Договор"].ToString() == dogovorcode)
                    {
                        listofoplata.Rows[i].Selected = true;
                        listofoplata.CurrentCell = listofoplata.Rows[i].Cells[2];
                    }
                    listofoplata.Rows[i].Cells[4].Value = reader["Подразделение"].ToString();
                    listofoplata.Rows[i].Cells[5].Value = Convert.ToDateTime(reader["ДатаДоговора"]);
                    listofoplata.Rows[i].Cells[6].Value = Convert.ToDateTime(reader["ДатаОкончания"]);
                    listofoplata.Rows[i].Cells[7].Value = Convert.ToDateTime(reader["ДатаВозврата"]);
                    listofoplata.Rows[i].Cells[8].Value = Convert.ToInt32(reader["СуммаДоговора"].ToString().Replace(" ", string.Empty));
                    dolg = reader["ДолгПоСумме"].ToString();
                    if (dolg.IndexOf(',') != -1) dolg = dolg.Remove(dolg.IndexOf(',')).Replace(" ", string.Empty);
                    else dolg = dolg.Replace(" ", string.Empty);
                    listofoplata.Rows[i].Cells[9].Value = Convert.ToInt32(dolg);
                    listofoplata.Rows[i].Cells[10].Value = Convert.ToInt32(reader["ДнейПросрочки"]);
                    if (reader["Пропуск"].ToString() != "")
                    {
                        listofoplata.Rows[i].Cells[11].Style.BackColor = Color.Green;
                        DateTime tasktime;
                        tasktime = Convert.ToDateTime(reader["Пропуск"].ToString());
                        int raznica = (DateTime.Now - tasktime).Days;
                        listofoplata.Rows[i].Cells[11].Value = raznica;
                    }
                    else listofoplata.Rows[i].Cells[11].Style.BackColor = Color.Gold;
                    DateTime datavozvr = Convert.ToDateTime(reader["ДатаВозврата"]);
                    int raznicad = (datavozvr - DateTime.Now).Days;

                    if (raznicad < 0 || Convert.ToInt32(reader["ДнейПросрочки"]) > 0)
                    {

                        listofoplata.Rows[i].Cells[7].Style.BackColor = Color.Red;
                    }
                    else if (raznicad == 0)
                    {
                        listofoplata.Rows[i].Cells[7].Style.BackColor = Color.Green;
                    }
                    else if (raznicad == 1)
                    {
                        listofoplata.Rows[i].Cells[7].Style.BackColor = Color.Orange;
                    }
                    listofoplata.Rows[i].Cells[12].Value = reader["notif"].ToString();
                    i++;
                }
                reader.Close();

            }
        }

        void dolg(string dogovorcode, NpgsqlCommand cmd)
        {


            string txt = "select Код, ФИОЗаемщика, Договор,Подразделение, СуммаДоговора,Пропуск,ДатаДоговора,ДатаОкончания,ДатаПослПлатежа,ДатаПослПропПлатежа,ДолгПоСумме,Просрочка,ОстатокПоДоговору,(select Тип from РаботаДолг where ФИО = ФИОЗаемщика order by ДатаРаботы desc Limit 1) as notif from Должники ";
            if (data.userrules == "Администратор" || data.userrules == "НСБ" || data.userrules == "Юрист")
            {
                txt += data.selectdolg + " order by Просрочка desc";
            }

            else
            {
                Console.WriteLine(data.usercitydop);
                String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                txt += "where (Подразделение = '" + data.usercity + "' ";
                for (int i = 0; i != words.Length; i++)
                {
                    txt += " or Подразделение = '" + words[i] + "'";
                }
                txt += " ) order by Просрочка desc";
            }
            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }

            visible(listofdoljniki, q);
            cmd.CommandText = txt;
            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    if (Convert.ToInt32(reader["Просрочка"]) > 0)
                    {
                        string dolg = "";
                        listofdoljniki.Rows.Add();
                        listofdoljniki.Rows[i].Cells[0].Value = i + 1;
                        listofdoljniki.Rows[i].Cells[1].Value = reader["Код"].ToString();
                        listofdoljniki.Rows[i].Cells[2].Value = reader["Договор"].ToString();
                        using (StreamReader fs = new StreamReader("\\\\" + data.ipfiles + "\\программа\\update\\dead.txt"))
                        {
                            while (true)
                            {

                                string temp = fs.ReadLine();

                                if (reader["Договор"].ToString() == temp)
                                {
                                    listofdoljniki.Rows[i].Cells[3].Style.BackColor = Color.Red;
                                }
                                if (temp == null) break;

                            }
                        }
                        if (reader["Подразделение"].ToString() == "Аскино")
                        {

                            using (StreamReader fs = new StreamReader("\\\\" + data.ipfiles + "\\программа\\update\\ugdela.txt"))
                            {
                                while (true)
                                {

                                    string temp = fs.ReadLine();

                                    if (reader["Договор"].ToString() == temp)
                                    {
                                        listofdoljniki.Rows[i].Cells[2].Style.BackColor = Color.Red;
                                    }
                                    if (temp == null) break;

                                }
                            }


                        }

                        listofdoljniki.Rows[i].Cells[3].Value = reader["ФИОЗаемщика"].ToString();
                        if (reader["Договор"].ToString() == dogovorcode)
                        {
                            listofdoljniki.Rows[i].Selected = true;
                            listofdoljniki.CurrentCell = listofdoljniki.Rows[i].Cells[2];
                        }
                        listofdoljniki.Rows[i].Cells[4].Value = reader["Подразделение"].ToString();
                        listofdoljniki.Rows[i].Cells[5].Value = Convert.ToDateTime(reader["ДатаДоговора"]);
                        listofdoljniki.Rows[i].Cells[6].Value = Convert.ToDateTime(reader["ДатаОкончания"]);
                        listofdoljniki.Rows[i].Cells[7].Value = Convert.ToDateTime(reader["ДатаПослПлатежа"]);
                        listofdoljniki.Rows[i].Cells[8].Value = Convert.ToDateTime(reader["ДатаПослПропПлатежа"]);
                        listofdoljniki.Rows[i].Cells[9].Value = Convert.ToInt32(reader["СуммаДоговора"].ToString().Replace(" ", string.Empty));

                        dolg = reader["ДолгПоСумме"].ToString();
                        dolg = dolg.Replace(" ", string.Empty);
                        listofdoljniki.Rows[i].Cells[10].Value = Convert.ToDouble(dolg);


                        listofdoljniki.Rows[i].Cells[11].Value = Convert.ToInt32(reader["Просрочка"]);
                        if (reader["Пропуск"].ToString() != "")
                        {
                            listofdoljniki.Rows[i].Cells[12].Style.BackColor = Color.Green;
                            DateTime tasktime;
                            tasktime = Convert.ToDateTime(reader["Пропуск"].ToString());
                            int raznica = (DateTime.Now - tasktime).Days;
                            listofdoljniki.Rows[i].Cells[12].Value = raznica;
                        }
                        else listofdoljniki.Rows[i].Cells[12].Style.BackColor = Color.Gold;
                        listofdoljniki.Rows[i].Cells[13].Value = reader["notif"].ToString();
                    }
                    i++;
                }
                reader.Close();

            }
        }

        void pretenzii(string dogovorcode, NpgsqlCommand cmd, int variant)
        {

            string txt = "select Код, ФИОЗаемщика,Шаблон, Договор,Подразделение,Пропуск, ДатаРешения,ДолгСуд,Оплатил,ДолгКПК,ДатаПлатежа,СуммаПлатежа,(select Тип from РаботаДолг where ФИО = ФИОЗаемщика order by ДатаРаботы desc Limit 1) as notif from семьшесть ";
            if (data.userrules == "Администратор" || data.userrules == "НСБ" || data.userrules == "Юрист")
            {
                txt += data.selectpretenzii;
            }
            else if (data.userrules == "Бухгалтерия")
            {

                txt += " where семьшесть.Договор in (select Договор from РаботаДолг where (Тип = 'получено судебное решение' or Тип = 'получен судебный приказ') and семьшесть.Договор=РаботаДолг.Договор and семьшесть.ФИОЗаемщика=РаботаДолг.ФИО) and семьшесть.Шаблон = 'Иск'";
            }
            else
            {
                String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                txt += "where (Подразделение = '" + data.usercity + "'";
                for (int i = 0; i != words.Length; i++)
                {
                    txt += "or Подразделение = '" + words[i] + "'";
                }
                txt += ")";
            }


            if (variant == 1)
            {
                txt = "select Код, ФИОЗаемщика,Шаблон, Договор,Подразделение,Пропуск, ДатаРешения,ДолгСуд,Оплатил,ДолгКПК,ДатаПлатежа,СуммаПлатежа,(select Тип from РаботаДолг where ФИО = ФИОЗаемщика order by ДатаРаботы desc Limit 1) as notif from семьшесть where ПодразделениеДляСБ = 'Айгузин'";
            }
            //if (variant == 2)
            //{
            //    txt = "select Код, ФИОЗаемщика,Шаблон, Договор,Подразделение,Пропуск, ДатаРешения,ДолгСуд,Оплатил,ДолгКПК,ДатаПлатежа,СуммаПлатежа,(select Тип from РаботаДолг where ФИО = ФИОЗаемщика order by ДатаРаботы desc Limit 1) as notif from семьшесть where ПодразделениеДляСБ = 'Бакиров'";
            //}

            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }


            visible(listofsud, q);

            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    string dolg, oplata, dolgsud, poslplatej = "";
                    listofsud.Rows.Add();
                    listofsud.Rows[i].Cells[0].Value = (i + 1).ToString();
                    listofsud.Rows[i].Cells[1].Value = reader["Код"].ToString();

                    listofsud.Rows[i].Cells[2].Value = reader["Договор"].ToString();
                    if (reader["Подразделение"].ToString() == "Аскино")
                    {

                        using (StreamReader fs = new StreamReader("\\\\" + data.ipfiles + "\\программа\\update\\ugdela.txt"))
                        {
                            while (true)
                            {

                                string temp = fs.ReadLine();

                                if (reader["Договор"].ToString() == temp)
                                {
                                    listofsud.Rows[i].Cells[2].Style.BackColor = Color.Red;
                                }
                                if (temp == null) break;

                            }
                        }


                    }
                    using (StreamReader fs = new StreamReader("\\\\" + data.ipfiles + "\\программа\\update\\dead.txt"))
                    {
                        while (true)
                        {

                            string temp = fs.ReadLine();

                            if (reader["Договор"].ToString() == temp || reader["ФИОЗаемщика"].ToString() == temp)
                            {
                                listofsud.Rows[i].Cells[3].Style.BackColor = Color.Red;
                            }
                            if (temp == null) break;

                        }
                    }
                    listofsud.Rows[i].Cells[3].Value = reader["ФИОЗаемщика"].ToString();
                    if (reader["Договор"].ToString() == dogovorcode)
                    {
                        listofsud.Rows[i].Selected = true;
                        listofsud.CurrentCell = listofsud.Rows[i].Cells[2];
                    }
                    listofsud.Rows[i].Cells[4].Value = reader["Подразделение"].ToString();
                    listofsud.Rows[i].Cells[7].Value = Convert.ToDateTime(reader["ДатаРешения"]);
                    dolgsud = reader["ДолгСуд"].ToString();
                    dolg = reader["ДолгКПК"].ToString();
                    oplata = reader["Оплатил"].ToString();

                    if (reader["ДатаПлатежа"].ToString() == "")
                    {
                        listofsud.Rows[i].Cells[5].Value = new DateTime(2045, 7, 20);
                    }
                    else
                    {
                        listofsud.Rows[i].Cells[5].Value = Convert.ToDateTime(reader["ДатаПлатежа"]);
                    }

                    poslplatej = reader["СуммаПлатежа"].ToString();


                    if (dolgsud.IndexOf(',') != -1) dolgsud = dolgsud.Remove(dolgsud.IndexOf(',')).Replace(" ", string.Empty);
                    else dolgsud = dolgsud.Replace(" ", string.Empty);
                    listofsud.Rows[i].Cells[8].Value = Convert.ToInt32(dolgsud);

                    if (oplata.IndexOf(',') != -1) oplata = oplata.Remove(oplata.IndexOf(',')).Replace(" ", string.Empty);
                    else oplata = oplata.Replace(" ", string.Empty);
                    listofsud.Rows[i].Cells[9].Value = Convert.ToInt32(oplata);

                    if (dolg.IndexOf(',') != -1) dolg = dolg.Remove(dolg.IndexOf(',')).Replace(" ", string.Empty);
                    else dolg = dolg.Replace(" ", string.Empty);
                    listofsud.Rows[i].Cells[6].Value = Convert.ToInt32(dolg);

                    if (poslplatej.IndexOf(',') != -1) poslplatej = poslplatej.Remove(poslplatej.IndexOf(',')).Replace(" ", string.Empty);
                    else poslplatej = poslplatej.Replace(" ", string.Empty);
                    listofsud.Rows[i].Cells[10].Value = Convert.ToInt32(poslplatej);
                    listofsud.Rows[i].Cells[12].Value = reader["notif"].ToString();
                    if (reader["Пропуск"].ToString() != "")
                    {
                        listofsud.Rows[i].Cells[11].Style.BackColor = Color.Green;
                        DateTime tasktime;
                        tasktime = Convert.ToDateTime(reader["Пропуск"].ToString());
                        int raznica = (DateTime.Now - tasktime).Days;
                        listofsud.Rows[i].Cells[11].Value = raznica;
                    }
                    else listofsud.Rows[i].Cells[11].Style.BackColor = Color.Gold;
                    i++;
                }
                reader.Close();

            }
        }

        void pretenziidead(string dogovorcode, NpgsqlCommand cmd)
        {

            string txt = "select Код, ФИОЗаемщика,Договор,Подразделение,ДолгКПК,Пропуск,(select Тип from РаботаДолг where ФИО = ФИОЗаемщика order by ДатаРаботы desc Limit 1) as notif from УмершиеПретензии ";
            //if (data.userrules == "Администратор" || data.userrules == "НСБ" || data.userrules == "Юрист")
            //{
            //    txt += data.selectpretenzii;
            //}
            //else
            //{
            //    String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            //    txt += "where (Подразделение = '" + data.usercity + "'";
            //    for (int i = 0; i != words.Length; i++)
            //    {
            //        txt += "or Подразделение = '" + words[i] + "'";
            //    }
            //    txt += ")";
            //}

            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }


            visible(listofsuddead, q);

            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    string dolg, oplata, dolgsud, poslplatej = "";
                    listofsuddead.Rows.Add();
                    listofsuddead.Rows[i].Cells[0].Value = (i + 1).ToString();
                    listofsuddead.Rows[i].Cells[1].Value = reader["Код"].ToString();
                    listofsuddead.Rows[i].Cells[2].Value = reader["Договор"].ToString();
                    listofsuddead.Rows[i].Cells[3].Value = reader["ФИОЗаемщика"].ToString();
                    if (reader["Договор"].ToString() == dogovorcode)
                    {
                        listofsuddead.Rows[i].Selected = true;
                        listofsuddead.CurrentCell = listofsuddead.Rows[i].Cells[2];
                    }
                    listofsuddead.Rows[i].Cells[4].Value = reader["Подразделение"].ToString();

                    dolg = reader["ДолгКПК"].ToString();


                    if (dolg.IndexOf(',') != -1) dolg = dolg.Remove(dolg.IndexOf(',')).Replace(" ", string.Empty);
                    else dolg = dolg.Replace(" ", string.Empty);
                    listofsuddead.Rows[i].Cells[6].Value = Convert.ToInt32(dolg);


                    listofsuddead.Rows[i].Cells[12].Value = reader["notif"].ToString();

                    if (reader["Пропуск"].ToString() != "")
                    {
                        listofsuddead.Rows[i].Cells[11].Style.BackColor = Color.Green;
                        DateTime tasktime;
                        tasktime = Convert.ToDateTime(reader["Пропуск"].ToString());
                        int raznica = (DateTime.Now - tasktime).Days;
                        listofsuddead.Rows[i].Cells[11].Value = raznica;
                    }
                    else listofsuddead.Rows[i].Cells[11].Style.BackColor = Color.Gold;
                    i++;
                }
                reader.Close();

            }
        }

        void zayavki(string dogovorcode, NpgsqlCommand cmd)
        {
            listofzauavki.Focus();

            string txt = "SELECT ФИО,Код,Подразделение,Сумма,Телефон,Дата,Работа, (select Тип from РаботаДолг where cast(Заявки.Код as text) = РаботаДолг.Путь order by ДатаРаботы desc Limit 1) as notif FROM Заявки";
            if (data.userrules == "Администратор" || data.userrules == "ГлавМенеджер")
            {
                txt += " " + data.selectzayavka + " order by Дата desc";
            }
            else
            {
                String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                txt += " where Подразделение = '" + words[0] + "' order by Дата desc";
            }
            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }

            visible(listofzauavki, q);
            cmd.CommandText = txt;
            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    listofzauavki.Rows.Add();
                    listofzauavki.Rows[i].Cells[0].Value = i + 1;
                    listofzauavki.Rows[i].Cells[3].Value = reader["Подразделение"].ToString();
                    listofzauavki.Rows[i].Cells[1].Value = reader["Код"].ToString();
                    if (reader["ФИО"].ToString() == dogovorcode)
                    {
                        listofzauavki.Rows[i].Selected = true;
                        listofzauavki.CurrentCell = listofzauavki.Rows[i].Cells[2];
                    }
                    listofzauavki.Rows[i].Cells[2].Value = reader["ФИО"].ToString();
                    listofzauavki.Rows[i].Cells[6].Value = reader["Телефон"].ToString();
                    listofzauavki.Rows[i].Cells[5].Value = reader["Сумма"].ToString();
                    listofzauavki.Rows[i].Cells[4].Value = Convert.ToDateTime(reader["Дата"]);
                    if (reader["notif"].ToString() == "Оформление" || reader["notif"].ToString() == "Отказ") listofzauavki.Rows[i].Cells[7].Style.BackColor = Color.Green;
                    else listofzauavki.Rows[i].Cells[7].Style.BackColor = Color.Red;
                    i++;
                }
                reader.Close();

            }

        }

        void users(string dogovorcode, NpgsqlCommand cmd)
        {

            string txt = "SELECT ФИО,Код,Подразделение,Логин,Вход,Должность,ДопПодразделение,Положение FROM Пользователи order by Логин";
            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }

            visible(listofusers, q);
            cmd.CommandText = txt;

            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    listofusers.Rows.Add();
                    listofusers.Rows[i].Cells[0].Value = i + 1;
                    listofusers.Rows[i].Cells[1].Value = reader["Код"].ToString();
                    listofusers.Rows[i].Cells[2].Value = reader["Логин"].ToString();
                    listofusers.Rows[i].Cells[3].Value = reader["Подразделение"].ToString();
                    listofusers.Rows[i].Cells[4].Value = reader["ФИО"].ToString();
                    listofusers.Rows[i].Cells[5].Value = reader["Должность"].ToString();
                    listofusers.Rows[i].Cells[6].Value = reader["ДопПодразделение"].ToString();
                    if (reader["Логин"].ToString() == dogovorcode)
                    {
                        listofusers.Rows[i].Selected = true;
                        listofusers.CurrentCell = listofusers.Rows[i].Cells[2];
                    }
                    if (reader["Вход"].ToString() == "1") listofusers.Rows[i].Cells[7].Value = true;
                    else listofusers.Rows[i].Cells[7].Value = false;
                    listofusers.Rows[i].Cells[8].Value = reader["Положение"].ToString();
                    i++;
                }

                reader.Close();

            }
        }

        void msk(string dogovorcode, NpgsqlCommand cmd)
        {

            string txt = "SELECT ФИО,Код,Подразделение,Дата,Менеджер,Телефон FROM Консультации ";
            if (data.userrules == "Администратор")
            {
                txt += " order by Дата desc";
            }
            else txt += " where Подразделение = '" + data.usercity + "' order by Дата desc";

            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }

            visible(listofconsulation, q);
            cmd.CommandText = txt;

            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    listofconsulation.Rows.Add();
                    listofconsulation.Rows[i].Cells[0].Value = reader["Код"].ToString();
                    listofconsulation.Rows[i].Cells[1].Value = reader["ФИО"].ToString();
                    listofconsulation.Rows[i].Cells[2].Value = reader["Подразделение"].ToString();
                    listofconsulation.Rows[i].Cells[3].Value = reader["Дата"].ToString();
                    listofconsulation.Rows[i].Cells[4].Value = reader["Менеджер"].ToString();
                    listofconsulation.Rows[i].Cells[5].Value = reader["Телефон"].ToString();
                    if (reader["ФИО"].ToString() == dogovorcode)
                    {
                        listofconsulation.Rows[i].Selected = true;
                        listofconsulation.CurrentCell = listofconsulation.Rows[i].Cells[1];
                    }
                    i++;
                }
                reader.Close();

            }
        }

        void sberej(string dogovorcode, NpgsqlCommand cmd)
        {

            string txt = "SELECT Код,ФИОЗаемщика,Подразделение,Договор,СуммаДоговора,СуммаВыплаты,ДатаВыплаты,Выдан FROM Сбережения";
            if (data.userrules == "Администратор" || data.userrules == "Бухгалтерия")
            {
                txt += " order by ДатаВыплаты desc";
            }
            else
            {
                String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                txt += " where (Подразделение = '" + data.usercity + "'";
                for (int i = 0; i != words.Length; i++)
                {
                    txt += " or Подразделение = '" + words[i] + "'";
                }
                txt += ")";
                txt += " order by ДатаВыплаты desc";
            }

            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }

            visible(sbertable, q);
            cmd.CommandText = txt;

            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    sbertable.Rows.Add();
                    sbertable.Rows[i].Cells[0].Value = i + 1;
                    sbertable.Rows[i].Cells[1].Value = reader["Код"].ToString();
                    sbertable.Rows[i].Cells[2].Value = reader["ФИОЗаемщика"].ToString();
                    sbertable.Rows[i].Cells[3].Value = reader["Подразделение"].ToString();
                    sbertable.Rows[i].Cells[4].Value = reader["СуммаВыплаты"].ToString();
                    sbertable.Rows[i].Cells[5].Value = reader["ДатаВыплаты"].ToString();
                    sbertable.Rows[i].Cells[6].Value = reader["Выдан"].ToString();
                    i++;
                }
                reader.Close();


            }
        }

        void proverka(string dogovorcode, NpgsqlCommand cmd)
        {

            string txt = "SELECT Код,Подразделение,ТипПроверки,Дата,ФИОСотрудника,Проверяющий,Проблема,ОтветПроверяющего,ОтветМенеджера,ДатаОтветаПроверяющего,ДатаОтветаМенеджера FROM ПроверкаМенеджеров";
            if (data.userrules.Contains("Администратор") || data.userrules == "Бухгалтерия")
            {
                selectofcheck = "";
                for (int e = 0; e < selectcheck.Length; e++)
                {
                    if (selectcheck[e] != "")
                    {
                        selectofcheck += selectcheck[e];
                    }
                }
                txt += " where Подразделение <> 's' " + selectofcheck + " order by Дата desc";
                toolStripMenuItem8.Visible = true;
                показатьToolStripMenuItem2.Visible = true;
                выбратьСотрудникаToolStripMenuItem.Visible = true;
                выбратьПроверяющегоToolStripMenuItem.Visible = true;
            }
            else if (data.userrules == "НСБ" || data.userrules == "Юрист" || data.userrules == "ГлавМенеджер")
            {
                txt += " where Проверяющий = '" + data.userFIO + "' and (ОтветПроверяющего = 'Нет' or ОтветМенеджера = 'Нет') order by Дата desc";
                toolStripMenuItem8.Visible = true;
            }
            else txt += " where Подразделение = '" + data.usercity + "'  order by Дата desc";
            cmd.CommandText = txt;
            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {

                while (reader.Read())
                {
                    q++;
                }
            }

            visible(listofcheck, q);
            cmd.CommandText = txt;
            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    listofcheck.Rows.Add();
                    listofcheck.Rows[i].Cells[0].Value = reader["Код"].ToString();
                    listofcheck.Rows[i].Cells[1].Value = i + 1;
                    listofcheck.Rows[i].Cells[2].Value = reader["Подразделение"].ToString();
                    listofcheck.Rows[i].Cells[3].Value = reader["ТипПроверки"].ToString();
                    listofcheck.Rows[i].Cells[4].Value = reader["ФИОСотрудника"].ToString();
                    listofcheck.Rows[i].Cells[5].Value = reader["Проверяющий"].ToString();
                    listofcheck.Rows[i].Cells[6].Value = reader["Дата"].ToString();
                    listofcheck.Rows[i].Cells[7].Value = reader["Проблема"].ToString();
                    if (reader["ОтветПроверяющего"].ToString() == "Да")
                    {
                        TimeSpan span = Convert.ToDateTime(reader["ДатаОтветаПроверяющего"]) - Convert.ToDateTime(reader["Дата"]);
                        string relative = span.Days.ToString();
                        if (relative == "0")
                        {
                            listofcheck.Rows[i].Cells[8].Style.BackColor = Color.Green;
                        }
                        else
                        {
                            listofcheck.Rows[i].Cells[8].Style.BackColor = Color.Orange;
                        }
                        listofcheck.Rows[i].Cells[8].Value = relative;
                    }
                    else
                    {
                        listofcheck.Rows[i].Cells[8].Style.BackColor = Color.Red;
                    }
                    if (reader["ОтветМенеджера"].ToString() == "Да")
                    {

                        TimeSpan span1 = Convert.ToDateTime(reader["ДатаОтветаМенеджера"]) - Convert.ToDateTime(reader["Дата"]);
                        string relative1 = span1.Days.ToString();
                        if (relative1 == "0")
                        {
                            listofcheck.Rows[i].Cells[9].Style.BackColor = Color.Green;
                        }
                        else
                        {
                            listofcheck.Rows[i].Cells[9].Style.BackColor = Color.Orange;
                        }
                        listofcheck.Rows[i].Cells[9].Value = relative1;
                    }
                    else
                    {
                        listofcheck.Rows[i].Cells[9].Style.BackColor = Color.Red;
                    }
                    if (reader["Дата"].ToString() == dogovorcode)
                    {
                        listofcheck.Rows[i].Selected = true;
                        listofcheck.CurrentCell = listofcheck.Rows[i].Cells[6];
                    }
                    i++;
                }
                reader.Close();

            }
        }

        void oldbase(string dogovorcode, NpgsqlCommand cmd, string type)
        {
            toolStripComboBox1.Items.Clear();
            toolStripComboBox1.Items.AddRange(data.options.DO());
            string txt = "";
            if (type == "актив")
            {
                txt = "SELECT Контакт,Код,Подразделение,Контрагент,Номер,Работа FROM ОбзвонКонтакты where (Работа = 'Придет' or Работа ='' or Работа ='Просит перезвонить' or Работа='Занято' or Работа='Недоступен')";
            }
            else if (type == "архив")
            {
                txt = "SELECT Контакт,Код,Подразделение,Контрагент,Номер,Работа FROM ОбзвонКонтакты where (Работа = 'Заключил договор' or Работа='Отказ' or Работа='Есть активный заем')";
            }
            else if (type == "База")
            {
                txt = "SELECT Контакт,Код,Подразделение,Контрагент,Работа FROM Обзвон where МенеджерОбзвона = '" + data.usercity + "' and Контрагент not in (select ФИОЗаемщика from Должники) and Контрагент not in (select ФИОЗаемщика from семьшесть)";
            }
            else if (type == "ПроверкаКачества")
            {
                txt = "SELECT Контрагент,Код,Подразделение,Контакт,Сумма,Вопрос1 FROM КонтрольКачестваЗвонки where Дата >= '" + DateTime.Today.AddDays(-5) + "'";
            }




            

            if (type == "ПроверкаКачества")
            {
                //if (data.userrules == "Администратор" || data.userrules == "ГлавМенеджер")
                //{
                   
                //}
                //else
                //{
                //    txt += " where Дата >= '" + DateTime.Today.AddDays(-3) + "'";
                //}
                cmd.CommandText = txt + " order by Дата";
            }
            else if (type == "База")
            {
               
                cmd.CommandText = txt + " order by Контрагент";
            }
            else
            {
               
                if (data.userrules == "Администратор" || data.userrules == "ГлавМенеджер")
                {
                    txt += data.selectoldbase;
                }
                else
                {
                    txt += " and (Подразделение = '" + data.usercity + "' or Подразделение = '" + data.usercityobzvon + "') ";
                }
                cmd.CommandText = txt + " order by Контрагент";
            }


            int q = 0;
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    q++;
                }
            }

            visible(tableofoldbase, q);
            using (var reader = cmd.ExecuteReader())
            {
                int i = 0;
                while (reader.Read())
                {
                    tableofoldbase.Rows.Add();
                    tableofoldbase.Rows[i].Cells[0].Value = i + 1;
                    tableofoldbase.Rows[i].Cells[1].Value = reader["Код"].ToString();
                    tableofoldbase.Rows[i].Cells[2].Value = reader["Контакт"].ToString();
                    tableofoldbase.Rows[i].Cells[3].Value = reader["Контрагент"].ToString();
                    tableofoldbase.Rows[i].Cells[4].Value = reader["Подразделение"].ToString();

                    if (type == "ПроверкаКачества")
                    {
                        if (reader["Вопрос1"].ToString() != "") tableofoldbase.Rows[i].Cells[6].Style.BackColor = Color.Green;
                        else tableofoldbase.Rows[i].Cells[6].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        tableofoldbase.Rows[i].Cells[5].Value = reader["Работа"].ToString();
                    }


                    if (reader["Код"].ToString() == dogovorcode)
                    {
                        tableofoldbase.Rows[i].Selected = true;
                        tableofoldbase.CurrentCell = tableofoldbase.Rows[i].Cells[2];
                    }
                    i++;
                }
                reader.Close();
            }
        }


        private void update()
        {
            this.Text = "КПК ФинансистЪ. Пользователь:" + data.userrules + "-" + data.userFIO + ".";
            string dogovorcode = "";
            try
            {
                if (data.typeoflist == 9 || data.typeoflist == 11)
                {
                    dogovorcode = table.Rows[table.CurrentRow.Index].Cells[1].Value.ToString();
                }
                else if (data.typeoflist == 10)
                {
                    dogovorcode = table.Rows[table.CurrentRow.Index].Cells[6].Value.ToString();
                }

                else
                {
                    dogovorcode = table.Rows[table.CurrentRow.Index].Cells[2].Value.ToString();
                }
            }
            catch { }

            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    //if (data.typeoflist <= 2 && (data.userrules == "Администратор" || data.userrules == "Юрист" || data.userrules == "НСБ" || data.userFIO == "Гараева Оксана Константиновна"))
                    if (data.typeoflist <= 2 && (data.userrules == "Администратор" || data.userrules == "Юрист" || data.userrules == "НСБ"))
                    {
                        label1.Visible = true;
                        findfoottxt.Visible = true;
                        label2.Visible = true;
                        findfootpodr.Visible = true;
                        findok.Visible = true;
                        findcancel.Visible = true;
                    }
                    else if (data.typeoflist == 11 || data.typeoflist == 13)
                    {
                        label1.Visible = true;
                        findfoottxt.Visible = true;
                        label2.Visible = true;
                        findfootpodr.Visible = true;
                        findok.Visible = true;
                        findcancel.Visible = true;
                    }
                    else
                    {
                        label1.Visible = false;
                        findfoottxt.Visible = false;
                        label2.Visible = false;
                        findfootpodr.Visible = false;
                        findok.Visible = false;
                        findcancel.Visible = false;
                    }

                    if (data.typeoflist == 0)
                    {
                        oplata(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 1)
                    {
                        dolg(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 2)
                    {
                        pretenzii(dogovorcode, cmd, 0);
                        if (Environment.MachineName == "ПЛАНШЕТ")
                        {
                            listofsud.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
                            listofsud.RowTemplate.Height = 35;
                        }
                    }
                    if (data.typeoflist == 12)
                    {
                        pretenziidead(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 3)
                    {
                        zayavki(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 4)
                    {
                        sberej(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 5)
                    {
                        proverka(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 6)
                    {
                        msk(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 11)
                    {
                        oldbase(dogovorcode, cmd, "актив");
                    }
                    if (data.typeoflist == 13)
                    {
                        oldbase(dogovorcode, cmd, "архив");
                    }
                    if (data.typeoflist == 14)
                    {
                        oldbase(dogovorcode, cmd, "База");
                    }
                    if (data.typeoflist == 15)
                    {
                        oldbase(dogovorcode, cmd, "ПроверкаКачества");
                    }
                    if (data.typeoflist == 10)
                    {
                        users(dogovorcode, cmd);
                    }
                    if (data.typeoflist == 98)
                    {
                        pretenzii(dogovorcode, cmd, 1);
                    }
                    //if (data.typeoflist == 99)
                    //{
                    //    pretenzii(dogovorcode, cmd, 2);
                    //}



                    if (data.typeoflist == 100)
                    {
                        foreach (Control control in this.Controls)
                        {
                            if (control is DataGridView)
                            {
                                control.Visible = false;
                            }
                        }
                        webBrowser2.Visible = true;
                        webBrowser2.DocumentText = "<html><style> .head{background:#33FFCC;text-align:center;}.txt{text-align:center;color:white;} body{background:red} label{text-align:center;} p{border-style:double;margin:0;}</style><body><br>asfafvadfaevc<br>asfafvadfaevc<br>asfafvadfaevc</body></html>";
                        webBrowser2.Show();
                    }
                }
            }
            if (data.typeoflist == 8)
            {
                string[,] DOO;
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select count(Код) from ДОms";
                        object value = cmd.ExecuteScalar();
                        DOO = new string[Convert.ToInt32(value), 2];
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
                выборкаToolStripMenuItem.Visible = true;
                table = listofcalls;
                visible(listofcalls, 0);
                listofcalls.Rows.Clear();
                selectcall = "";
                for (int q = 0; q < selectcalls.Length; q++)
                {
                    if (selectcalls[q] != "")
                    {
                        selectcall += selectcalls[q];
                    }
                }

                using (var conn = new MySqlConnection(data.stringconnect()[0]))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand())
                    {
                        cmd.Connection = conn;
                        if (napravcalls == 0)
                        {
                            cmd.CommandText = "SELECT src,dst ,disposition , calldate,billsec,recordingfile,uniqueid FROM cdr where LENGTH(src) = 3 " + datecall + selectcall + " order by calldate";

                        }
                        else
                        {
                            cmd.CommandText = "SELECT src,dst ,disposition ,calldate,billsec,recordingfile,uniqueid FROM cdr where LENGTH(src) > 3 " + datecall + selectcall + " order by calldate";
                        }
                        using (var reader = cmd.ExecuteReader())
                        {
                            int i = 0;
                            while (reader.Read())
                            {
                                listofcalls.Rows.Add();
                                listofcalls.Rows[i].Cells[0].Value = i + 1;
                                listofcalls.Rows[i].Cells[2].Value = reader["uniqueid"].ToString();
                                for (int q = 0; q < DOO.GetLength(0); q++)
                                {
                                    if (DOO[q, 1] == reader["src"].ToString())
                                    {
                                        listofcalls.Rows[i].Cells[1].Value = DOO[q, 0];
                                        break;
                                    }
                                    else { listofcalls.Rows[i].Cells[1].Value = reader["src"].ToString(); }
                                }
                                for (int q = 0; q < DOO.GetLength(0); q++)
                                {
                                    if (DOO[q, 1] == reader["dst"].ToString())
                                    {
                                        listofcalls.Rows[i].Cells[3].Value = DOO[q, 0];
                                        break;
                                    }
                                    else { listofcalls.Rows[i].Cells[3].Value = reader["dst"].ToString(); ; }
                                }


                                listofcalls.Rows[i].Cells[4].Value = reader["calldate"].ToString();
                                listofcalls.Rows[i].Cells[5].Value = reader["billsec"].ToString();
                                listofcalls.Rows[i].Cells[6].Value = reader["disposition"].ToString();
                                listofcalls.Rows[i].Cells[7].Value = reader["recordingfile"].ToString();
                                listofcalls.Rows[i].Cells[8].Value = reader["src"].ToString();
                                listofcalls.Rows[i].Cells[9].Value = reader["dst"].ToString();

                                if (reader["uniqueid"].ToString() == dogovorcode)
                                {
                                    listofcalls.Rows[i].Selected = true;
                                    listofcalls.CurrentCell = listofcalls.Rows[i].Cells[1];
                                }
                                i++;
                            }
                        }
                    }
                }
            }
        }

        private void visible(DataGridView data1, int q)
        {
            foreach (Control control in this.Controls)
            {
                if (control is DataGridView)
                {
                    control.Visible = false;
                }
            }
            data1.Focus();
            table = data1;
            data1.Rows.Clear();
            data1.Visible = true;
            labelfoot.Text = "Всего: " + q.ToString();

        }

        private void findok1(object sender, EventArgs e)
        {
            fio = "LOWER(ФИОЗаемщика) like LOWER('" + findfoottxt.Text + "%')";
            if (data.typeoflist == 0)
            {
                if (podr != "") { data.selectoplata = "and " + fio + " and " + podr; }
                else { data.selectoplata = "and " + fio; }
            }
            if (data.typeoflist == 1)
            {
                if (podr != "") { data.selectdolg = "where " + fio + " and " + podr; }
                else { data.selectdolg = "where " + fio; }
            }
            if (data.typeoflist == 2)
            {
                if (podr != "") { data.selectpretenzii = "where " + fio + " and " + podr; }
                else { data.selectpretenzii = "where " + fio; }
            }
            if (data.typeoflist == 11 || data.typeoflist == 13)
            {
                obzvonmanager = 1;
                fio = "LOWER(ФИО) like LOWER('" + findfoottxt.Text + "%')";
                if (podr != "") { data.selectoldbase = "and " + fio + " and " + podr; }
                else { data.selectoldbase = "and " + fio; }
            }
            update();
        }

        private void findcancel1(object sender, EventArgs e)
        {
            data.selectoplata = "";
            data.selectdolg = "";
            data.selectpretenzii = "";
            data.selectoldbase = "";
            findfoottxt.Text = "";
            findfootpodr.SelectedIndex = 0;
            fio = "";
            podr = "";
            update();
            obzvonmanager = 0;
        }

        private void podrfind(object sender, EventArgs e)
        {
            if (findfootpodr.Text != "")
            {
                podr = "Подразделение = '" + findfootpodr.Text + "'";
            }
            else
            {
                podr = "Подразделение <> '" + findfootpodr.Text + "'";
            }
            if (data.typeoflist == 0)
            {
                if (fio != "") { data.selectoplata = "and " + fio + " and " + podr; }
                else { data.selectoplata = "and " + podr; }
            }
            if (data.typeoflist == 1)
            {
                if (fio != "") { data.selectdolg = "where " + fio + " and " + podr; }
                else { data.selectdolg = "where " + podr; }
            }
            if (data.typeoflist == 2)
            {
                if (fio != "") { data.selectpretenzii = "where " + fio + " and " + podr; }
                else { data.selectpretenzii = "where " + podr; }
            }
            if (data.typeoflist == 11 || data.typeoflist == 13)
            {
                obzvonmanager = 1;
                if (fio != "") { data.selectoldbase = "and " + fio + " and " + podr; }
                else { data.selectoldbase = "and " + podr; }
            }
            update();
            table.Focus();
        }


        private void colorbtn(ToolStripMenuItem btn)
        {
            data.selectoplata = "";
            data.selectdolg = "";
            data.selectpretenzii = "";
            findfoottxt.Text = "";
            findfootpodr.SelectedIndex = 0;
            fio = "";
            podr = "";
            refreshpages = "0";
            double offset = 0;
            double Limit = ((this.Height - (27 + 34 + 22 + 44)) / 22);


            limit = " offset " + offset + " limit " + Limit + "";
            for (int i = 0; i < menuStrip1.Items.Count; i++)
            {
                if (menuStrip1.Items[i].Name.Contains("btn"))
                {
                    menuStrip1.Items[i].BackColor = Color.DarkCyan;
                }
            }
            btn.BackColor = Color.Cyan;
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            parol.TextBox.PasswordChar = '*';

            username.Items.Clear();
            using (var conn = new NpgsqlConnection(data.path))
            {
                AsterNET.Manager.ManagerConnection manager = new ManagerConnection("ip", port, "login", "pass");
                manager.Login();
                manager.NewState += new NewStateEventHandler(Monitoring_NewState);
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select Логин from Пользователи order by Логин";
                    cmd.ExecuteNonQuery();
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            username.Items.Add(reader["Логин"].ToString());
                        }
                    }
                }
                conn.Close();
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.Visible = true;
                notifyIcon1.ShowBalloonTip(5, "Оповещение", "Программа свернулась. Нажмите на иконку, чтобы снова развернуть программу.", ToolTipIcon.Info);
            }
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            if (data.updatetable)
            {
                if (data.active == 1)
                {
                    update();
                }
                data.updatetable = false;
            }
        }

        static void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Suspend)
            {
                Application.Exit();
            }
        }

        void vxod()
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "Select Пароль, Код, Телефон, ДопПодразделение, Подразделение, Должность, ФИО, Права,Роль,Показ From Пользователи Where Логин = '" + username.Text + "'";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["Пароль"].ToString().ToUpper() == parol.Text.ToUpper())
                            {
                                data.usercityobzvon = reader["Права"].ToString();
                                data.usercode = reader["Код"].ToString();
                                data.userphone = reader["Телефон"].ToString();
                                data.usercity = reader["Подразделение"].ToString();
                                data.usercitydop = reader["ДопПодразделение"].ToString();
                                data.userrules = reader["Роль"].ToString();
                                data.userpermission = reader["Показ"].ToString();
                                data.userFIO = reader["ФИО"].ToString();
                                //data.usercityobzvon = reader["ПодразделениеДляОбзвона"].ToString();
                                data.username = username.Text;
                                String[] userpermission = data.userpermission.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                                for (int i = 0; i != userpermission.Length; i++)
                                {
                                    if (userpermission[i] == "0")
                                    {
                                        dogovorabtn.Visible = true;
                                    }
                                    if (userpermission[i] == "1")
                                    {
                                        doljnikibtn.Visible = true;
                                    }
                                    if (userpermission[i] == "2")
                                    {
                                        sevensixschetbtn.Visible = true;
                                    }
                                    if (userpermission[i] == "3")
                                    {
                                        zayavkabtn.Visible = true;
                                    }
                                    if (userpermission[i] == "4")
                                    {
                                        sberbtn.Visible = true;
                                    }
                                    if (userpermission[i] == "5")
                                    {
                                        Checkingbtn.Visible = true;
                                    }
                                    if (userpermission[i] == "6")
                                    {
                                        konsbtn.Visible = true;
                                    }
                                    if (userpermission[i] == "7")
                                    {
                                        phonebtn.Visible = true;
                                        обзвонToolStripMenuItem.Visible = true;
                                    }
                                    if (userpermission[i] == "99")
                                    {
                                        bir1btn.Visible = true;
                                    }
                                    //if (userpermission[i] == "98")
                                    //{
                                    //    bir2btn.Visible = true;
                                    //}
                                    if (userpermission[i] == "12")
                                    {
                                        теплаяБазаToolStripMenuItem.Visible = true;
                                    }
                                    if (userpermission[i] == "13")
                                    {
                                        проверкаКачестваОбслуживанияToolStripMenuItem.Visible = true;
                                    }
                                }

                                statusStrip1.Items.Clear();

                                statusStrip1.Items.Add(labelfoot);
                                findcancel.BackColor = Color.White;
                                findcancel.Text = "Отмена";
                                findcancel.Click += this.findcancel1;
                                label1.Text = "ФИО  ";
                                findfoottxt.BackColor = Color.White;
                                label2.Text = "Подразделение  ";
                                findfootpodr.BackColor = Color.White;
                                findfootpodr.DropDownStyle = ComboBoxStyle.DropDownList;
                                findfootpodr.Items.Clear();
                                findfootpodr.Items.Add("");
                                findfootpodr.Items.AddRange(data.options.DOactual());
                                findfootpodr.SelectedIndexChanged += this.podrfind;
                                findfoottxt.TextChanged += this.findok1;

                                data.typeoflist = Convert.ToInt32(userpermission[0]);
                                actionsbtn.Visible = true;
                                refresh_button.Visible = true;

                                if (data.userrules == "Администратор" || data.userrules == "ГлавМенеджер" || data.userrules == "НСБ" || data.userrules == "Юрист")
                                {
                                    statusStrip1.Items.Add(label1);
                                    statusStrip1.Items.Add(findfoottxt);
                                    statusStrip1.Items.Add(label2);
                                    statusStrip1.Items.Add(findfootpodr);
                                    statusStrip1.Items.Add(findcancel);
                                }
                                if (data.userrules == "Администратор")
                                {
                                    analiz_button.Visible = true;
                                    analiz.Visible = true;
                                    звонкиToolStripMenuItem1.Visible = true;
                                    users_button.Visible = true;
                                    actionsbtn.Visible = true;
                                    phonebtn.Visible = true;
                                    звонкиToolStripMenuItem1.Visible = true;
                                    colorbtn(dogovorabtn);
                                    table = listofoplata;
                                }
                                else if (data.userrules == "Менеджер" || data.userrules == "ГлавМенеджер")
                                {
                                    colorbtn(dogovorabtn);
                                    table = listofoplata;
                                }
                                else if (data.userrules == "СБ" || data.userrules == "НСБ" || data.userrules == "Юрист")
                                {
                                    colorbtn(doljnikibtn);
                                    table = listofdoljniki;
                                }
                                else if (data.userrules == "Бухгалтерия")
                                {
                                    colorbtn(dogovorabtn);
                                    table = listofsud;
                                    analiz_button.Visible = true;
                                }
                                using (StreamWriter sw = new StreamWriter(@"C:\КПК\Программа\login.ini"))
                                {
                                    sw.WriteLine(data.username);
                                }
                                label4.Text = "";
                                statusStrip1.Items.Add(label4);
                                toolStripMenuItem1.Visible = false;
                                username.Visible = false;
                                парольToolStripMenuItem.Visible = false;
                                parol.Visible = false;
                                enter_button.Visible = false;
                                //data.typeoflist = 100;
                                update();
                                data.active = 1;
                                //webBrowser2.Visible = true;
                                //webBrowser2.DocumentText = "<html><style> .head{background:#33FFCC;text-align:center;}.txt{text-align:center;color:white;} body{background:red} label{text-align:center;} p{border-style:double;margin:0;}</style><body><br>asfafvadfaevc<br>asfafvadfaevc<br>asfafvadfaevc</body></html>";
                                //webBrowser2.Show();
                            }
                            else { MessageBox.Show("Неверный пароль"); parol.Text = ""; data.active = 0; }
                        }
                        reader.Close();
                    }
                }
            }
            //string txtdahboard = "На сегодня у вас, в :"+data.usercity+Environment.NewLine;
            //using (var conn = new NpgsqlConnection(data.path))
            //{
            //    conn.Open();
            //    using (var cmd = new NpgsqlCommand())
            //    {
            //        cmd.Connection = conn;
            //        cmd.CommandText = "select СуммаДолг ,КолВоДолг,КолВоПрет,СуммаПрет from АнализСБ where Подразделение = '" + data.usercity + "' and ДатаФорм ='"+DateTime.Today+"'";
            //        using (var reader = cmd.ExecuteReader())
            //        {
            //            while (reader.Read())
            //            {
            //                txtdahboard += "текущие долги: " + reader["СуммаДолг"].ToString() + " рублей и количество: " + reader["КолВоДолг"].ToString()+Environment.NewLine;
            //                txtdahboard += "текущие долги по 76 счету: " + reader["СуммаПрет"].ToString() + " рублей и количество: " + reader["КолВоПрет"].ToString() + Environment.NewLine;
            //            }
            //        }
            //    }
            //}
            //richTextBox1.Text = txtdahboard;
            if (data.active == 1)
            {
                timer2.Enabled = true;
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "update Пользователи set Вход ='1', Положение ='" + DateTime.Now + "' where Логин = '" + data.username + "'";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into ОтслежВходВыход (Пользователь,Время,Действие) values ('" + data.username + "', '" + DateTime.Now + "', 'Вход')";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Select Обновить From Пользователи Where Логин = '" + data.username + "'";
                        object upd = cmd.ExecuteScalar();
                        if (upd.ToString() == "1")
                        {
                            MessageBox.Show("Программа сейчас обновится, нажмите 'ОК' и подождите");
                            System.Diagnostics.Process.Start("\\\\" + data.ipfiles + "\\программа\\update\\update.exe");
                            Application.Exit();
                        }
                    }
                }
                //refreshnotif();
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "update Пользователи set Положение ='" + DateTime.Now + "' where Логин = '" + data.username + "'";
                    cmd.ExecuteNonQuery();
                }
            }
            refreshnotif();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (data.username != "")
            {
                using (var conn = new NpgsqlConnection(data.path))
                {

                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "update Пользователи set Вход ='0', Положение ='" + DateTime.Now + "' where Логин = '" + data.username + "'";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into ОтслежВходВыход (Пользователь,Время,Действие) values ('" + data.username + "', '" + DateTime.Now + "', 'Выход')";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser1.ShowPrintPreviewDialog();
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Maximized;
            notifyIcon1.Visible = false;
        }


        void refreshnotif()
        {
            data.notiftext = "";
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;

                    if (data.userrules == "Администратор")
                    {
                        notificationurist(cmd, 0);
                        notificationurist(cmd, 1);
                        notificationurist(cmd, 2);
                        notificationurist(cmd, 3);
                        notificationurist(cmd, 4);
                        notificationurist(cmd, 5);
                        notificationurist(cmd, 6);
                        notificationurist(cmd, 7);
                        notificationurist(cmd, 8);
                    }
                    if (data.userrules == "Бухгалтерия")
                    {
                        notificationsud(cmd);
                    }
                    if (data.userrules == "Менеджер")
                    {
                        notificationsber(cmd);
                        notificationdogovora(cmd, "cast(ДнейПросрочки as int) > '0' and cast(ДнейПросрочки as int) < '3' and", 1);
                        notificationdogovora(cmd, "cast(ДнейПросрочки as int) > '2' and", 2);
                        notificationzayavki(cmd);
                        notificationchecking(cmd);
                    }
                    if (data.userrules == "СБ")
                    {
                        string query = "Подразделение = '" + data.usercity + "'";
                        String[] words = data.usercitydop.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i != words.Length; i++)
                        {
                            query += " or Подразделение = '" + words[i] + "'";
                        }
                        notificationsb(cmd, query, 0);
                        notificationsb(cmd, query, 2);
                        notificationsb(cmd, query, 1);
                        notificationsb(cmd, query, 3);
                        notificationsb(cmd, query, 4);
                        notificationsb(cmd, query, 5);
                        notificationsb(cmd, query, 6);

                    }
                    if (data.userrules == "Юрист" || data.userrules == "НСБ")
                    {
                        notificationurist(cmd, 0);
                        notificationurist(cmd, 1);
                        notificationurist(cmd, 2);
                        notificationurist(cmd, 3);
                        notificationurist(cmd, 4);
                        notificationurist(cmd, 5);
                        notificationurist(cmd, 6);
                        notificationurist(cmd, 7);
                        notificationurist(cmd, 8);
                    }

                }
            }
            if (data.notiftext != "")
            {
                new notification().Show();
            }
        }

        void notificationdogovora(NpgsqlCommand cmd, string days, int q)
        {
            cmd.CommandText = "select ФИОЗаемщика from Оплата where Оплата.Договор not in (select Договор from РаботаДолг where Тип <> '' and Оплата.Договор = РаботаДолг.Договор and Дата = '" + DateTime.Now.ToString("dd.MM.yyyy") + "') and " + days + " Подразделение = '" + data.usercity + "'";
            Console.WriteLine(cmd.CommandText);
            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    if (q == 1)
                    {
                        data.notiftext += "<div class=head><p><b>Заемщик вышел в просрочку!</b></p></div>";
                    }
                    if (q == 2)
                    {
                        data.notiftext += "<div class=head><p><b>Просрочка по договору составляет более 3 дней</b></p></div>";
                    }
                    while (reader.Read())
                    {
                        data.notiftext += "<div class=txt><label>" + reader["ФИОЗаемщика"].ToString() + "</label></div>";
                    }
                }
            }
        }

        void notificationdogovoraadmin(NpgsqlCommand cmd, string days, int q)
        {
            cmd.CommandText = "select ФИОЗаемщика from Оплата where Оплата.Договор not in (select Договор from РаботаДолг where Тип <> '' and Оплата.Договор = РаботаДолг.Договор and Дата = '" + DateTime.Now.ToString("dd.MM.yyyy") + "') and " + days + " Подразделение = '" + data.usercity + "'";
            Console.WriteLine(cmd.CommandText);
            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    if (q == 1)
                    {
                        data.notiftext += "<div class=head><p><b>Заемщик вышел в просрочку!</b></p></div>";
                    }
                    if (q == 2)
                    {
                        data.notiftext += "<div class=head><p><b>Просрочка по договору составляет более 3 дней</b></p></div>";
                    }
                    while (reader.Read())
                    {
                        data.notiftext += "<div class=txt><label>" + reader["ФИОЗаемщика"].ToString() + "</label></div>";
                    }
                }
            }
        }

        void notificationsber(NpgsqlCommand cmd)
        {
            cmd.CommandText = "select ФИОЗаемщика from Сбережения where Выдан = 'Нет' and Подразделение = '" + data.usercity + "'";
            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    data.notiftext += "<div class=head><p><b>Должна быть выплата по вкладам</b></p></div>";
                    while (reader.Read())
                    {
                        data.notiftext += "<div class=txt><label>" + reader["ФИОЗаемщика"].ToString() + "</label></div>";
                    }
                }
            }
        }

        void notificationsb(NpgsqlCommand cmd, string podr, int q)
        {
            var txt = "";
            if (q == 0)
            {

                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'отмена судебного решения' and (cast(now() as date) - cast(ДатаРаботы as date)) > 10 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'исковое заявление' or Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'Банкротство' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нужно подать исковое после отмены</b></p></div>";
            }
            if (q == 1)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'получено судебное решение' and (cast(now() as date) - cast(ДатаРаботы as date)) > 30 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'заявление в ССП' or Тип = 'получен исполнительный лист' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нужно получить исполнительный лист</b></p></div>";
            }
            if (q == 2)
            {
                cmd.CommandText = "select ФИОЗаемщика,Шаблон,Подразделение,Договор from Должники where Должники.Договор not in (select Договор from РаботаДолг where (Тип = 'претензия' or Тип = 'заявление на судебный приказ' or Тип = 'исковое заявление' or Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'получен судебный приказ' or Тип = 'отмена судебного решения' or Тип = 'заявление в ССП') and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and cast(ДнейПросрочки as int)< 30 and cast(ДнейПросрочки as int)>= 10 and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нет претензии</b></p></div>";
            }
            if (q == 3)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор not in (select Договор from РаботаДолг where (Тип = 'заявление на судебный приказ' or Тип = 'исковое заявление' or Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'получен судебный приказ' or Тип = 'отмена судебного решения' or Тип = 'заявление в ССП' or Тип = 'Банкротство') and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор in (select Договор from РаботаДолг where Тип = 'претензия' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and cast(ДнейПросрочки as int)> 60 and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>В суд</b></p></div>";
            }
            if (q == 4)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'заявление на судебный приказ' and (cast(now() as date) - cast(ДатаРаботы as date)) > 30 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'отмена судебного решения' or Тип = 'получен судебный приказ' or Тип = 'Банкротство' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нужно получить судебный приказ</b></p></div>";
            }
            if (q == 5)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'исковое заявление' and (cast(now() as date) - cast(ДатаРаботы as date)) > 45 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'получено судебное решение' or Тип = 'Банкротство' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>нужно получить решение суда</b></p></div>";
            }
            if (q == 6)
            {
                cmd.CommandText = "select ФИОЗаемщика,Шаблон,Подразделение,Договор from семьшесть where семьшесть.Договор not in (select Договор from РаботаДолг where Тип <> '' and РаботаДолг.Договор=семьшесть.Договор and (cast(now() as date) - cast(ДатаРаботы as date)) > -30) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нет работы более 30 дней по 76 счету</b></p></div>";
            }


            if (q == 7)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'получен исполнительный лист' or Тип = 'получен судебный приказ' and (cast(now() as date) - cast(ДатаРаботы as date)) > 10 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'заявление в ССП' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нужно писать заявление в ССП</b></p></div>";
            }
            if (q == 8)
            {
                cmd.CommandText = "select ФИОЗаемщика,Шаблон,Подразделение,Договор from семьшесть where семьшесть.Договор in (select Договор from РаботаДолг where Тип = 'ИП прекращено' and РаботаДолг.Договор=семьшесть.Договор and (cast(now() as date) - cast(ДатаРаботы as date)) >= 180) and (" + podr + ") order by ФИОЗаемщика";
                txt = "<div class=head><p><b>Нужно по-новой заявление в ССП</b></p></div>";
            }


            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    data.notiftext += txt;
                    while (reader.Read())
                    {
                        var j = 0;
                        if (reader["Подразделение"].ToString() == "Аскино")
                        {
                            using (StreamReader fs = new StreamReader("\\\\" + data.ipfiles + "\\программа\\update\\ugdela.txt"))
                            {
                                while (true)
                                {
                                    string temp = fs.ReadLine();

                                    if (reader["Договор"].ToString() == temp)
                                    {
                                        j = 1;
                                    }
                                    if (temp == null) break;
                                }
                            }
                        }
                        if (j == 0)
                        {
                            data.notiftext += "<div class=txt><label>" + reader["ФИОЗаемщика"].ToString() + "</label></div>";
                        }
                    }
                }
            }
        }

        void notificationurist(NpgsqlCommand cmd, int q)
        {
            var txt = "";
            if (q == 0)
            {
                cmd.CommandText = "select ФИОЗаемщика,Шаблон,Подразделение,Договор from Должники where Должники.Договор not in (select Договор from РаботаДолг where (Тип = 'претензия' or Тип = 'заявление на судебный приказ' or Тип = 'исковое заявление' or Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'получен судебный приказ' or Тип = 'отмена судебного решения' or Тип = 'заявление в ССП') and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and cast(ДнейПросрочки as int)> 20 order by Подразделение";
                txt = "<div class=head><p><b>Нет претензии</b></p></div>";
            }
            if (q == 1)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор not in (select Договор from РаботаДолг where (Тип = 'заявление на судебный приказ' or Тип = 'исковое заявление' or Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'получен судебный приказ' or Тип = 'отмена судебного решения' or Тип = 'заявление в ССП' or Тип = 'Банкротство') and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор in (select Договор from РаботаДолг where Тип = 'претензия' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and cast(ДнейПросрочки as int)> 60 order by Подразделение";
                txt = "<div class=head><p><b>В суд</b></p></div>";
            }
            if (q == 2)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'отмена судебного решения' and (cast(now() as date) - cast(ДатаРаботы as date)) > 10 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'исковое заявление' or Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'Банкротство' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) order by Подразделение";
                txt = "<div class=head><p><b>Нужно подать исковое после отмены</b></p></div>";
            }

            if (q == 3)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'заявление на судебный приказ' and (cast(now() as date) - cast(ДатаРаботы as date)) > 30 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'отмена судебного решения' or Тип = 'получен судебный приказ' or Тип = 'Банкротство' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) order by Подразделение";
                txt = "<div class=head><p><b>Нужно получить судебный приказ</b></p></div>";
            }
            if (q == 4)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'исковое заявление' and (cast(now() as date) - cast(ДатаРаботы as date)) > 45 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'получено судебное решение' or Тип = 'получен исполнительный лист' or Тип = 'получено судебное решение' or Тип = 'Банкротство' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) order by Подразделение";
                txt = "<div class=head><p><b>Нужно получить решение суда</b></p></div>";
            }

            if (q == 5)
            {
                cmd.CommandText = "select ФИОЗаемщика,Шаблон,Подразделение,Договор from семьшесть where семьшесть.Договор not in (select Договор from РаботаДолг where Тип <> '' and РаботаДолг.Договор=семьшесть.Договор and (cast(now() as date) - cast(ДатаРаботы as date)) > 60) order by Подразделение";
                txt = "<div class=head><p><b>Нет работы более 60 дней по 76 счету</b></p></div>";
            }
            if (q == 6)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'получено судебное решение' and (cast(now() as date) - cast(ДатаРаботы as date)) > 30 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'заявление в ССП' or Тип = 'получен исполнительный лист' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) order by Подразделение";
                txt = "<div class=head><p><b>Нужно получить исполнительный лист</b></p></div>";
            }
            if (q == 7)
            {
                cmd.CommandText = "select ФИОЗаемщика,Подразделение,Договор from Должники where Должники.Договор in (select Договор from РаботаДолг where Тип = 'получен исполнительный лист' or Тип = 'получен судебный приказ' and (cast(now() as date) - cast(ДатаРаботы as date)) > 10 and РаботаДолг.ФИО=Должники.ФИОЗаемщика) and Должники.Договор not in (select Договор from РаботаДолг where Тип = 'заявление в ССП' and РаботаДолг.ФИО=Должники.ФИОЗаемщика) order by Подразделение";
                txt = "<div class=head><p><b>Нужно писать заявление в ССП</b></p></div>";
            }
            if (q == 8)
            {
                cmd.CommandText = " select ФИОЗаемщика,Шаблон,Подразделение,Договор from семьшесть where семьшесть.Договор in (select Договор from РаботаДолг where Тип = 'ИП прекращено' and РаботаДолг.Договор=семьшесть.Договор and (cast(now() as date) - cast(ДатаРаботы as date)) >= 180) order by Подразделение";
                txt = "<div class=head><p><b>Нужно по-новой заявление в ССП</b></p></div>";
            }


            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    data.notiftext += txt;
                    while (reader.Read())
                    {
                        var j = 0;
                        if (reader["Подразделение"].ToString() == "Аскино")
                        {
                            using (StreamReader fs = new StreamReader("\\\\" + data.ipfiles + "\\программа\\update\\ugdela.txt"))
                            {
                                while (true)
                                {
                                    string temp = fs.ReadLine();

                                    if (reader["Договор"].ToString() == temp)
                                    {
                                        j = 1;
                                    }
                                    if (temp == null) break;
                                }
                            }
                        }
                        if (j == 0)
                        {
                            data.notiftext += "<div class=txt><label>" + reader["ФИОЗаемщика"].ToString() + "   " + reader["Подразделение"].ToString() + "</label></div>";
                        }
                    }
                }
            }
        }

        void notificationchecking(NpgsqlCommand cmd)
        {
            cmd.CommandText = "SELECT Код,Подразделение,ТипПроверки,Дата,ФИОСотрудника,Проверяющий,Проблема,ОтветПроверяющего,ОтветМенеджера,ДатаОтветаПроверяющего,ДатаОтветаМенеджера FROM ПроверкаМенеджеров where Подразделение = '" + data.usercity + "' and ФИОСотрудника = '" + data.userFIO + "' and ОтветМенеджера = 'Нет'";
            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    data.notiftext += "<div class=head><p><b>Не исправлена ошибка</b></p></div>";
                    while (reader.Read())
                    {
                        data.notiftext += "<div class=txt>от <label>" + reader["Проверяющий"].ToString() + "</label></div>";
                        data.notiftext += "<div class=txt>по <label>" + reader["Проблема"].ToString() + "</label></div>";
                    }
                }
            }
        }

        void notificationsud(NpgsqlCommand cmd)
        {
            cmd.CommandText = "select Код, ФИОЗаемщика, Договор from семьшесть where семьшесть.Договор in (select Договор from РаботаДолг where (Тип = 'получено судебное решение' or Тип = 'получен судебный приказ') and семьшесть.Договор=РаботаДолг.Договор and семьшесть.ФИОЗаемщика=РаботаДолг.ФИО) and семьшесть.Шаблон = 'Иск'";


            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    data.notiftext += "<div class=head><p><b>Можно оформлять решение суда:</b></p></div>";
                    while (reader.Read())
                    {
                        data.notiftext += "<div class=txt><label>" + reader["ФИОЗаемщика"].ToString() + "</label></div>";
                    }
                }
            }
        }

        void notificationzayavki(NpgsqlCommand cmd)
        {
            object count = null;
            try
            {
                using (var connmy = new MySqlConnection(data.stringconnect()[5]))
                {
                    connmy.Open();
                    using (var cmdmy = new MySqlCommand())
                    {
                        cmdmy.Connection = connmy;
                        cmdmy.CommandText = "SELECT count(*) FROM zayavka";
                        count = cmdmy.ExecuteScalar();

                    }
                }
            }
            catch { }
            if (Convert.ToInt32(count) != 0)
            {
                string[,] value = new string[Convert.ToInt32(count), 5];
                using (var connmy = new MySqlConnection(data.stringconnect()[5]))
                {
                    connmy.Open();
                    using (var cmdmy = new MySqlCommand())
                    {
                        cmdmy.Connection = connmy;
                        cmdmy.CommandText = "SELECT FIO,code,phone,summa,date,city FROM zayavka";
                        using (MySqlDataReader reader = cmdmy.ExecuteReader())
                        {
                            int i = 0;
                            while (reader.Read())
                            {
                                value[i, 0] = reader["FIO"].ToString();
                                value[i, 1] = reader["city"].ToString();
                                value[i, 2] = reader["date"].ToString();
                                value[i, 3] = reader["summa"].ToString();
                                value[i, 4] = reader["phone"].ToString();
                                i++;
                            }
                        }
                        cmdmy.CommandText = "delete FROM zayavka";
                        cmdmy.ExecuteNonQuery();
                    }
                }
                using (var connps = new NpgsqlConnection(data.path))
                {
                    connps.Open();
                    using (var cmdps = new NpgsqlCommand())
                    {
                        cmdps.Connection = connps;
                        for (int i = 0; i != Convert.ToInt32(count); i++)
                        {
                            cmdps.CommandText = "insert into Заявки (ФИО,Подразделение,Дата,Сумма,Телефон) Values ('" + value[i, 0] + "','" + value[i, 1] + "','" + value[i, 2] + "','" + value[i, 3] + "','" + value[i, 4] + "')";
                            cmdps.ExecuteNonQuery();
                        }
                    }
                }
            }
            cmd.CommandText = "select ФИО, (select Тип from РаботаДолг where cast(Заявки.Код as text) = РаботаДолг.Путь order by ДатаРаботы desc Limit 1) as раб from Заявки where Дата>'" + new DateTime(2017, 11, 01) + "' and Подразделение = '" + data.usercitydop + "' and (((select Тип from РаботаДолг where cast(Заявки.Код as text) = РаботаДолг.Путь order by ДатаРаботы desc Limit 1)<>'Оформление' and (select Тип from РаботаДолг where cast(Заявки.Код as text) = РаботаДолг.Путь order by ДатаРаботы desc Limit 1)<>'Отказ') or (select Тип from РаботаДолг where cast(Заявки.Код as text) = РаботаДолг.Путь order by ДатаРаботы desc Limit 1) is null)";
            using (var reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    data.notiftext += "<div class=head><p><b>Не проведена конечная работа по заявке от</b></p></div>";
                    while (reader.Read())
                    {
                        data.notiftext += "<div class=txt><label>" + reader["ФИО"].ToString() + "</label></div>";
                    }
                }
            }



        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            vxod();
        }

        private void настрйокиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            update();
            refreshnotif();
        }

        private void menuStrip1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                обновитьToolStripMenuItem.Visible = true;
            }
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "update Пользователи set Обновить = '1' Where Логин <> 'Администратор'";
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        private void должникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(doljnikibtn);
            table = listofdoljniki;
            data.typeoflist = 1;
            update();
        }

        private void счетToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void договораToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            colorbtn(dogovorabtn);
            table = listofoplata;
            data.typeoflist = 0;
            update();
        }

        private void качествоToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            new analitics().ShowDialog();
        }

        private void toolStripTextBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                vxod();
            }
        }

        private void подключениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(zayavkabtn);
            object count = null;
            using (var conn = new MySqlConnection(data.stringconnect()[5]))
            {
                conn.Open();
                using (var cmd = new MySqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "SELECT count(*) FROM zayavka";
                    count = cmd.ExecuteScalar();

                }
            }
            if (Convert.ToInt32(count) != 0)
            {
                string[,] value = new string[Convert.ToInt32(count), 5];
                using (var conn = new MySqlConnection(data.stringconnect()[5]))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT FIO,code,phone,summa,date,city FROM zayavka";
                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            int i = 0;
                            while (reader.Read())
                            {
                                value[i, 0] = reader["FIO"].ToString();
                                value[i, 1] = reader["city"].ToString();
                                value[i, 2] = reader["date"].ToString();
                                value[i, 3] = reader["summa"].ToString();
                                value[i, 4] = reader["phone"].ToString();
                                i++;
                            }
                        }
                        cmd.CommandText = "delete FROM zayavka";
                        cmd.ExecuteNonQuery();
                    }
                }
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        for (int i = 0; i != Convert.ToInt32(count); i++)
                        {
                            cmd.CommandText = "insert into Заявки (ФИО,Подразделение,Дата,Сумма,Телефон) Values ('" + value[i, 0] + "','" + value[i, 1] + "','" + value[i, 2] + "','" + value[i, 3] + "','" + value[i, 4] + "')";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                table = listofzauavki;
                data.typeoflist = 3;
                update();
            }
            else
            {
                table = listofzauavki;
                data.typeoflist = 3;
                update();
            }
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sb = new StringBuilder("<html><head><meta charset='utf-8'><style>table {border-collapse: collapse;}th {border: 1px solid black; padding: 2px;} td {border: 1px solid black;padding: 2px;} </style>")
            .Append("</head><body><table><tr>");
            foreach (DataGridViewColumn c in table.Columns)
                sb.Append("<th>").Append(c.HeaderText).Append("</th>");
            foreach (DataGridViewRow o in table.Rows)
            {
                sb.Append("<tr>");
                foreach (DataGridViewCell i in o.Cells)
                    if (i.Value != null) sb.Append("<td>").Append(i.Value.ToString()).Append("</td>");
                    else sb.Append("<td>").Append("").Append("</td>");
                sb.Append("</tr>");
            }
            sb.Append("</table></body></html>").ToString();
            StreamWriter streamwriter = new StreamWriter(@"C:\КПК\Программа\index.html");
            streamwriter.WriteLine(sb);
            streamwriter.Close();
            webBrowser1.DocumentText = System.IO.File.ReadAllText(@"C:\КПК\Программа\index.html");

        }

        private void пользователиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(actionsbtn);
            table = listofusers;
            data.typeoflist = 10;
            update();
        }

        private void выборкаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            new select().ShowDialog();
        }

        private void анализToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new analitics().Show();
        }

        private void позвонитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AsterNET.Manager.ManagerConnection manager = new ManagerConnection(data.stringconnect()[1], Convert.ToInt32(data.stringconnect()[2]), data.stringconnect()[3], data.stringconnect()[4]);
            manager.Login();
            manager.SendAction((ManagerAction)new OriginateAction()
            {
                Channel = "SIP/" + data.userphone,
                CallerId = data.userphone,
                Context = "from-internal",
                Exten = listofcalls.CurrentCell.Value.ToString(),
                Priority = "1",
                Timeout = 20000,
                Async = true
            }, 50000);
        }

        private void воспроизвестиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string txt = listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[4].Value.ToString();
            string audio = "";
            audio = "\\\\***\\calls\\" + txt.Substring(6, 4) + "\\" + txt.Substring(3, 2) + "\\" + txt.Substring(0, 2) + "\\";
            audio += listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[7].Value.ToString();
            if (audio != "")
            {
               
                System.Diagnostics.Process.Start(audio);
            }
            else { MessageBox.Show("нету аудио"); }
        }

        private void toolStripMenuItem2_Click_1(object sender, EventArgs e)
        {
            data.adiingtype = "cons";
            new adding().ShowDialog();
        }

        private void консультацииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(konsbtn);
            table = listofconsulation;
            data.typeoflist = 6;
            update();
        }

        private void исходящиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
            table = listofcalls;
            data.typeoflist = 8;
            napravcalls = 0;
            update();
        }

        private void входящиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
            table = listofcalls;
            data.typeoflist = 8;
            napravcalls = 1;
            update();
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            colorbtn(Checkingbtn);
            table = listofcheck;
            data.typeoflist = 5;
            update();
        }

        private void toolStripMenuItem8_Click_1(object sender, EventArgs e)
        {
            data.checkaction = "add";
            new select().ShowDialog();
        }





        private void окToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sotr = listofcheck.Rows[listofcheck.CurrentRow.Index].Cells[4].Value.ToString();
            selectcheck[1] = " and ФИОСотрудника = '" + sotr + "' ";
            окToolStripMenuItem.Visible = false;
            отменаToolStripMenuItem2.Visible = true;
            update();

        }

        private void всеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            selectcheck[0] = "";
            update();
        }

        private void проверенныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            selectcheck[0] = " and (ОтветПроверяющего = 'Да' or ОтветПроверяющего = 'Да')";
            update();
        }

        private void неПроверенныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            selectcheck[0] = " and (ОтветПроверяющего = 'Нет' or ОтветПроверяющего = 'Нет')";
            update();

        }

        private void отменаToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            окToolStripMenuItem.Visible = true;
            отменаToolStripMenuItem2.Visible = false;
            selectcheck[1] = "";
            update();
        }

        private void окToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string sotr = listofcheck.Rows[listofcheck.CurrentRow.Index].Cells[5].Value.ToString();
            selectcheck[2] = " and Проверяющий = '" + sotr + "' ";
            окToolStripMenuItem1.Visible = false;
            отменаToolStripMenuItem3.Visible = true;
            update();
        }

        private void отменаToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            окToolStripMenuItem1.Visible = true;
            отменаToolStripMenuItem3.Visible = false;
            selectcheck[2] = "";
            update();
        }

        private void окToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            string datemysqlfield = toolStripTextBox1.Text.Substring(4, 4) + "-" + toolStripTextBox1.Text.Substring(2, 2) + "-" + toolStripTextBox1.Text.Substring(0, 2);
            datecall = " and LEFT(CAST(calldate as char), 10) = '" + datemysqlfield + "'";
            update();
        }

        private void окToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            отменаToolStripMenuItem1.Visible = true;
            окToolStripMenuItem2.Visible = false;
            selectcalls[0] = " and src = '" + listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[8].Value.ToString() + "'";
            update();
        }

        private void отменаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            отменаToolStripMenuItem1.Visible = false;
            окToolStripMenuItem2.Visible = true;
            selectcalls[0] = "";
            update();
        }

        private void окToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            окToolStripMenuItem3.Visible = false;
            отменаToolStripMenuItem4.Visible = true;
            selectcalls[1] = " and dst = '" + listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[9].Value.ToString() + "'"; ;
            update();
        }

        private void отменаToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            окToolStripMenuItem3.Visible = true
                ;
            отменаToolStripMenuItem4.Visible = false;
            selectcalls[1] = "";
            update();
        }

        private void всеToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            selectcalls[2] = "";
            update();
        }

        private void отвеченныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            selectcalls[2] = " and disposition = 'ANSWERED'";
            update();
        }

        private void analiz_Click(object sender, EventArgs e)
        {
            data.whatanaliz = "телефония";
            new analiz().ShowDialog();
        }

        private void аналитикаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            new analitics().ShowDialog();
        }

        private void фотоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("\\\\" + data.ipfiles + "\\программа\\update\\takeaphoto\\takephoto.exe");
        }

        private void показатьОперацииПользователяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            data.usercode = listofusers.Rows[listofusers.CurrentRow.Index].Cells[2].Value.ToString();
            new usersspy().ShowDialog();
        }

        private void обзвонToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


        private void dataGridView2_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                try
                {
                    data.dogovortablecode = listofsud.Rows[listofsud.CurrentRow.Index].Cells[1].Value.ToString();
                    new dogovor().ShowDialog();

                }
                catch { MessageBox.Show("Выберите договор"); }
            }
            if (e.Button == System.Windows.Forms.MouseButtons.Right) { }
        }

        private void dataGridView3_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                try
                {
                    data.dogovortablecode = listofzauavki.Rows[listofzauavki.CurrentRow.Index].Cells[1].Value.ToString();
                    new dogovor().ShowDialog();
                }
                catch { MessageBox.Show("Выберите заявку"); }
            }

        }

        private void listofusers_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == System.Windows.Forms.MouseButtons.Left)
            //{
            //    try
            //    {
            //        data.dogovortablecode = listofusers.Rows[listofusers.CurrentRow.Index].Cells[1].Value.ToString();
            //        new dogovor().ShowDialog();
            //    }
            //    catch { MessageBox.Show("Выберите пользоваетля"); }
            //}
        }

        private void listofoplata_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                try
                {
                    data.dogovortablecode = listofoplata.Rows[listofoplata.CurrentRow.Index].Cells[1].Value.ToString();
                    new dogovor().ShowDialog();
                }
                catch { MessageBox.Show("Выберите договор"); }
            }
        }

        private void listofdoljniki_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                try
                {
                    data.dogovortablecode = listofdoljniki.Rows[listofdoljniki.CurrentRow.Index].Cells[1].Value.ToString();
                    new dogovor().ShowDialog();
                }
                catch { MessageBox.Show("Выберите договор"); }
            }
        }

        private void tableofoldbase_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                
                    if (data.typeoflist == 15)
                    {
                        data.dogovortablecode = tableofoldbase.Rows[tableofoldbase.CurrentRow.Index].Cells[1].Value.ToString();

                    }
                    else 
                    {
                        data.dogovortablecode = tableofoldbase.Rows[tableofoldbase.CurrentRow.Index].Cells[2].Value.ToString();
                    }
                    new dogovor().ShowDialog();
                
            }
        }

        private void listofcalls_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 7)
            {
                string txt = listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[4].Value.ToString();
                string audio = "";
                audio = "\\\\***\calls\\" + txt.Substring(6, 4) + "\\" + txt.Substring(3, 2) + "\\" + txt.Substring(0, 2) + "\\";
                audio += listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[7].Value.ToString();
                if (audio != "")
                {
                 
                    System.Diagnostics.Process.Start(audio);
                }
                else { MessageBox.Show("нету аудио"); }
            }
        }

        private void listofconsulation_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                data.dogovortablecode = listofconsulation.Rows[listofconsulation.CurrentRow.Index].Cells[0].Value.ToString();
                new dogovor().ShowDialog();
            }
            catch { MessageBox.Show("Выберите клиента"); }
        }

        private void listofcheck_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            data.checkaction = "edit";
            data.checkcode = listofcheck.Rows[e.RowIndex].Cells[0].Value.ToString();
            new select().ShowDialog();
            //if (e.ColumnIndex == 8)
            //{
            //    using (var conn = new NpgsqlConnection(data.path))
            //    {
            //        conn.Open();
            //        using (var cmd = new NpgsqlCommand())
            //        {
            //            cmd.Connection = conn;
            //            cmd.CommandText = "select ДатаОтветаПроверяющего from ПроверкаМенеджеров where Код = '" + listofcheck.Rows[e.RowIndex].Cells[0].Value + "'";
            //            using (var reader = cmd.ExecuteReader())
            //            {
            //                while (reader.Read())
            //                {
            //                    MessageBox.Show(reader["ДатаОтветаПроверяющего"].ToString());
            //                }
            //                reader.Close();
            //            }
            //        }
            //    }
            //}
            //if (e.ColumnIndex == 9)
            //{
            //    using (var conn = new NpgsqlConnection(data.path))
            //    {
            //        conn.Open();
            //        using (var cmd = new NpgsqlCommand())
            //        {
            //            cmd.Connection = conn;
            //            cmd.CommandText = "select ДатаОтветаМенеджера from ПроверкаМенеджеров where Код = '" + listofcheck.Rows[e.RowIndex].Cells[0].Value + "'";
            //            using (var reader = cmd.ExecuteReader())
            //            {
            //                while (reader.Read())
            //                {
            //                    MessageBox.Show(reader["ДатаОтветаМенеджера"].ToString());
            //                }
            //                reader.Close();
            //            }
            //        }
            //    }
            //}
            //if (e.ColumnIndex == 7)
            //{
            //    MessageBox.Show(listofcheck.Rows[e.RowIndex].Cells[7].Value.ToString());
            //}
        }

        private void входящиеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
        }

        private void исходящиеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
        }

        private void окToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            selectoldbase[0] = "and Подразделение = '" + toolStripComboBox1.Text + "'";
            update();
        }

        private void отменаToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            selectoldbase[0] = "";
            update();

        }

        private void начатыеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            selectoldbase[1] = " and Работа <> ''";
            update();
        }

        private void неВРаботеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            selectoldbase[1] = " and Работа is null";
            update();
        }

        private void всеToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            selectoldbase[1] = "";
            update();
        }











        private void sberbtn_Click(object sender, EventArgs e)
        {
            colorbtn(sberbtn);
            table = sbertable;
            data.typeoflist = 4;
            update();
        }



        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            data.adiingtype = "sber";
            new adding().ShowDialog();

        }

        private void выданоToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string code = sbertable.Rows[sbertable.CurrentRow.Index].Cells[1].Value.ToString();
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "update Сбережения set Выдан = 'Да' where Код = '" + code + "'";
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    update();
                }
            }
        }

        private void переоформленоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string code = sbertable.Rows[sbertable.CurrentRow.Index].Cells[1].Value.ToString();
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "update Сбережения set Выдан = 'переоформление' where Код = '" + code + "'";
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    update();
                }
            }
        }

        private void listofsud_Sorted(object sender, EventArgs e)
        {
            for (int i = 0; i < listofsud.RowCount; i++)
            {
                listofsud.Rows[i].Cells[0].Value = i + 1;
            }

        }

        private void listofoplata_Sorted(object sender, EventArgs e)
        {
            for (int i = 0; i < listofoplata.RowCount; i++)
            {
                listofoplata.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void listofdoljniki_Sorted(object sender, EventArgs e)
        {
            for (int i = 0; i < listofdoljniki.RowCount; i++)
            {
                listofdoljniki.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void listofsud_SelectionChanged(object sender, EventArgs e)
        {

            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                listofsud.SelectionMode = DataGridViewSelectionMode.FullRowSelect; sum = 0;
            }
            else
            {
                listofsud.SelectionMode = DataGridViewSelectionMode.CellSelect;
                try
                {
                    if (listofsud.SelectedCells.Count == 1)
                    {
                        sum = Convert.ToInt32(listofsud.Rows[listofsud.CurrentRow.Index].Cells[listofsud.CurrentCell.ColumnIndex].Value);
                        label4.Text = "Сумма: " + sum.ToString();
                    }
                    if (listofsud.SelectedCells.Count > 1)
                    {
                        sum += Convert.ToInt32(listofsud.Rows[listofsud.CurrentRow.Index].Cells[listofsud.CurrentCell.ColumnIndex].Value);
                        label4.Text = "Сумма: " + sum.ToString();
                    }
                }
                catch { }
            }
        }

        private void listofoplata_SelectionChanged(object sender, EventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                listofoplata.SelectionMode = DataGridViewSelectionMode.FullRowSelect; sum = 0;
            }
            else
            {
                listofoplata.SelectionMode = DataGridViewSelectionMode.CellSelect;
                try
                {
                    if (listofoplata.SelectedCells.Count == 1)
                    {
                        sum = Convert.ToInt32(listofoplata.Rows[listofoplata.CurrentRow.Index].Cells[listofoplata.CurrentCell.ColumnIndex].Value);
                        label4.Text = "Сумма: " + sum.ToString();
                    }
                    if (listofoplata.SelectedCells.Count > 1)
                    {
                        sum += Convert.ToInt32(listofoplata.Rows[listofoplata.CurrentRow.Index].Cells[listofoplata.CurrentCell.ColumnIndex].Value);
                        label4.Text = "Сумма: " + sum.ToString();
                    }
                }
                catch { }
            }
        }

        private void listofdoljniki_SelectionChanged(object sender, EventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                listofdoljniki.SelectionMode = DataGridViewSelectionMode.FullRowSelect; sum = 0;
            }
            else
            {
                listofdoljniki.SelectionMode = DataGridViewSelectionMode.CellSelect;
                try
                {
                    if (listofdoljniki.SelectedCells.Count == 1)
                    {
                        sum = Convert.ToInt32(listofdoljniki.Rows[listofdoljniki.CurrentRow.Index].Cells[listofdoljniki.CurrentCell.ColumnIndex].Value);
                        label4.Text = "Сумма: " + sum.ToString();
                    }
                    if (listofdoljniki.SelectedCells.Count > 1)
                    {
                        sum += Convert.ToInt32(listofdoljniki.Rows[listofdoljniki.CurrentRow.Index].Cells[listofdoljniki.CurrentCell.ColumnIndex].Value);
                        label4.Text = "Сумма: " + sum.ToString();
                    }
                }
                catch { }
            }
        }

        private void parol_Click(object sender, EventArgs e)
        {
            if (Environment.MachineName == "ПЛАНШЕТ")
            {
                System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Windows) + "\\system32\\osk.exe");
                parol.Focus();
            }
        }

        private void toExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataObject dataObj = listofsud.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        

        private void bir1btn_Click(object sender, EventArgs e)
        {
            colorbtn(bir1btn);
            table = listofsud;
            data.typeoflist = 98;
            update();
        }

        private void bir2btn_Click(object sender, EventArgs e)
        {
            colorbtn(bir2btn);
            table = listofsud;
            data.typeoflist = 99;
            update();
        }

        private void счетToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            colorbtn(sevensixschetbtn);
            table = listofsud;
            data.typeoflist = 2;
            update();
        }

        private void чсетУмершиеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(sevensixschetbtn);
            table = listofsuddead;
            data.typeoflist = 12;
            update();
        }

        private void listofsuddead_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                try
                {
                    data.dogovortablecode = listofsuddead.Rows[listofsuddead.CurrentRow.Index].Cells[1].Value.ToString();
                    new dogovor().ShowDialog();

                }
                catch { MessageBox.Show("Выберите договор"); }
            }
            if (e.Button == System.Windows.Forms.MouseButtons.Right) { }
        }

        private void активныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
            table = tableofoldbase;
            data.typeoflist = 11;
            update();
        }

        private void архивToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
            table = tableofoldbase;
            data.typeoflist = 13;
            update();
        }

        private void теплаяБазаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
            table = tableofoldbase;
            data.typeoflist = 14;
            update();
        }


        void Monitoring_NewState(object sender, NewStateEvent e)
        {
            string state = e.ChannelStateDesc;
            string callerID = e.CallerIdNum;
            //Console.WriteLine(state);
            //Console.WriteLine(callerID);

            if ((state == "Ringing"))
            {
                //Console.WriteLine(callerID);
                //if (callerID != "user2")
                //{
                    data.managernum = callerID;
                    Console.WriteLine("~~~~~~~~~~~~~~~номер менеджера: {0}", data.managernum);
                //}
            }
            if ((state == "Ring"))
            {
                data.numberin = callerID;
                Console.WriteLine("~~~~~~~~~~~~~~~номер звонящего: {0}", data.numberin);
                //Console.WriteLine(callerID);
            }


            if (callerID != "user2" && data.managernum == data.userphone && state == "Ringing")
            {

                string strmes = "Звонит ";
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "select ФИО,Подразделение from НомераТелефонов where Номер='" + data.numberin + "'";
                        Console.WriteLine("~~~~~~~~~~~~~~~" + cmd.CommandText);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int i = 0;
                                while (reader.Read())
                                {
                                    if (i >= 1)
                                    {
                                        strmes += " или ";
                                    }
                                    strmes += reader["ФИО"].ToString() + " из " + reader["Подразделение"].ToString();

                                    i++;
                                }
                            }
                            else { strmes = "Звонит неизвестный номер"; }
                        }
                        

                    }
                }
                MessageBox.Show(strmes);
            }

        }

        private void проверкаКачестваОбслуживанияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorbtn(phonebtn);
            table = tableofoldbase;
            data.typeoflist = 15;
            update();
        }











    }
}
