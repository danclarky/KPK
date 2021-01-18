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
    public partial class analitics : Form
    {
        public analitics()
        {
            InitializeComponent();
        }

        public string[] clientofoplata, clientofoplatadone, clientofoplatanotwork;

        private void analitics_Load(object sender, EventArgs e)
        {

        }
        private void tableofanalizworkmanagers_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void button1_Click(object sender, EventArgs e)
        {
            int q = 0;
            if (comboBox1.SelectedIndex == 0)
            {
                TimeSpan oneDay = TimeSpan.FromDays(1);
                oplatadays.Rows.Clear();
                oplatadays.Columns.Clear();
                oplatadays.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Дата", HeaderText = "Дата", Width = 70 });
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    oplatadays.Columns.Add(new DataGridViewTextBoxColumn() { Name = data.DOanal()[q], HeaderText = data.DOanal()[q] + " План|Зависло|Оплата", Width = 200 });
                }
                int w = 0;
                for (DateTime date = datestart.Value; date <= dateend.Value; date += oneDay)
                {
                    oplatadays.Rows.Add();
                    oplatadays.Rows[w].Cells[0].Value = date.ToShortDateString();
                    w++;
                }
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    w = 0;
                    for (DateTime date = datestart.Value; date <= dateend.Value; date += oneDay)
                    {
                        int count = 0, count1 = 0, coun = 0;
                        double summa = 0, summa1 = 0, sum = 0;
                        using (var conn = new NpgsqlConnection(data.path))
                        {
                            conn.Open();
                            using (var cmd = new NpgsqlCommand())
                            {
                                cmd.Connection = conn;
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствНачало where Подразделение = '" + data.DOanal()[q] + "' and ДатаОплаты = '" + date + "'";
                                using (var reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            if (reader["summa"].ToString() != "") summa = Convert.ToDouble(reader["summa"]);
                                            else summa = 0;
                                            if (reader["kol"].ToString() != "") count = Convert.ToInt32(reader["kol"]);
                                            else count = 0;
                                            if (summa == 0) oplatadays.Rows[w].Cells[q + 1].Value = summa + "р; " + count.ToString() + " | ";
                                            else oplatadays.Rows[w].Cells[q + 1].Value = summa.ToString("#") + "р; " + count.ToString() + " | ";
                                        }
                                    }
                                }
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствТекущееНакопление where Подразделение = '" + data.DOanal()[q] + "' and ДатаОплаты = '" + date + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                                using (var reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            if (reader["summa"].ToString() != "") sum = Convert.ToDouble(reader["summa"]);
                                            else sum = 0;
                                            if (reader["kol"].ToString() != "") coun = Convert.ToInt32(reader["kol"]);
                                            else coun = 0;
                                            summa1 = summa - sum;
                                            count1 = count - coun;

                                            if (sum == 0) oplatadays.Rows[w].Cells[q + 1].Value += sum + "р; " + coun.ToString() + " | ";
                                            else oplatadays.Rows[w].Cells[q + 1].Value += sum.ToString("#") + "р; " + coun.ToString() + " | ";

                                            if (summa1 == 0) oplatadays.Rows[w].Cells[q + 1].Value += summa1 + "р;" + count1.ToString();
                                            else oplatadays.Rows[w].Cells[q + 1].Value += summa1.ToString("#") + "р;" + count1.ToString();
                                        }
                                    }
                                }
                            }
                        }
                        w++;
                    }
                }

            }
            if (comboBox1.SelectedIndex == 1)
            {
                int count = 0, count1 = 0, coun = 0;
                double summa = 0, summa1 = 0, sum = 0;
                oplataperiod.Rows.Clear();
                int[] oplataplankol = new int[data.DOanal().Length];
                double[] oplataplansum = new double[data.DOanal().Length];
                int[] oplatarealkol = new int[data.DOanal().Length];
                double[] oplatarealsum = new double[data.DOanal().Length];
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствНачало where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "'";
                            //}
                            //else
                            {
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствНачало where Подразделение = '" + data.DOanal()[q] + "' and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplataplankol[q] = Convert.ToInt32(reader["kol"]);
                                            count += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplataplansum[q] = Convert.ToDouble(reader["summa"]);
                                            summa += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствТекущееНакопление where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                            //}
                            //else 
                            {
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствТекущееНакопление where Подразделение = '" + data.DOanal()[q] + "' and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplatarealkol[q] = Convert.ToInt32(reader["kol"]);
                                            count1 += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplatarealsum[q] = Convert.ToDouble(reader["summa"]);
                                            summa1 += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }

                        }
                    }
                    coun = oplataplankol[q] - oplatarealkol[q];
                    sum = oplataplansum[q] - oplatarealsum[q];

                    oplataperiod.Rows.Add();
                    oplataperiod.Rows[q].Cells[0].Value = data.DOanal()[q];

                    if (sum == 0) oplataperiod.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum;
                    else oplataperiod.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum.ToString("#.#");
                    if (oplatarealsum[q] == 0) oplataperiod.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q];
                    else oplataperiod.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q].ToString("#.#");
                    if (oplataplansum[q] == 0) oplataperiod.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q];
                    else oplataperiod.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q].ToString("#.#");

                    double kach = 0, kachsum = 0;
                    if (oplataplankol[q] != 0)
                        kach = (double)coun / oplataplankol[q] * 100;
                    if (coun == 0)
                        kach = 0;
                    if (oplataplansum[q] != 0)
                        kachsum = (double)sum / oplataplansum[q] * 100;
                    if (sum == 0)
                        kachsum = 0;
                    oplataperiod.Rows[q].Cells[4].Value = Math.Round(kach, 0).ToString() + "%  | " + Math.Round(kachsum, 0).ToString() + "%";
                }
                oplataperiod.Rows.Add();
                oplataperiod.Rows[q].Cells[0].Value = "Всего";
                oplataperiod.Rows[q].Cells[1].Value = count.ToString() + " | " + summa.ToString("#.#");
                oplataperiod.Rows[q].Cells[2].Value = count1.ToString() + " | " + summa1.ToString("#.#");
                oplataperiod.Rows[q].Cells[3].Value = (count - count1).ToString() + " | " + (summa - summa1).ToString("#.#");
                double kach1 = 0, kachsum1 = 0;
                if (count != 0)
                    kach1 = (double)(count - count1) / count * 100;
                if (summa != 0)
                    kachsum1 = (double)(summa - summa1) / summa * 100;

                oplataperiod.Rows[q].Cells[4].Value = Math.Round(kach1, 0).ToString() + "%  | " + Math.Round(kachsum1, 0).ToString() + "%";
            }
          
            if (comboBox1.SelectedIndex == 2)
            {
                

                int[] oplataplankol = new int[data.template().Length];
                double[] oplataplansum = new double[data.template().Length];
                int[] oplatarealkol = new int[data.template().Length];
                double[] oplatarealsum = new double[data.template().Length];
                oplataperiodtemplate.Rows.Clear();
                oplataperiodtemplate.Columns.Clear();
                oplataperiodtemplate.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Шаблон", HeaderText = "Шаблон", Width = 150 });
                oplataperiodtemplate.Columns[0].Frozen = true;
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    oplataperiodtemplate.Columns.Add(new DataGridViewTextBoxColumn() { Name = data.DOanal()[q], HeaderText = data.DOanal()[q] + " План|Зависло|Оплата", Width = 200 });
                }
                for (q = 0; q < data.template().Length; q++)
                {
                    oplataperiodtemplate.Rows.Add();
                    oplataperiodtemplate.Rows[q].Cells[0].Value = data.template()[q];
                }
                oplataperiodtemplate.Rows.Add();
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    int count = 0, count1 = 0, coun = 0, totalkol = 0, totalkol1 = 0, totalkol2 = 0;
                    double summa = 0, summa1 = 0, sum = 0, totalsum = 0, totalsum1 = 0, totalsum2 = 0;
                    for (int w = 0; w < data.template().Length; w++)
                    {
                        using (var conn = new NpgsqlConnection(data.path))
                        {
                            conn.Open();
                            using (var cmd = new NpgsqlCommand())
                            {
                                cmd.Connection = conn;
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствНачало where Шаблон = '" + data.template()[w] + "' and Подразделение = '" + data.DOanal()[q] + "' and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "'";
                                using (var reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            if (reader["summa"].ToString() != "") summa = Convert.ToDouble(reader["summa"]);
                                            else summa = 0;
                                            if (reader["kol"].ToString() != "") count = Convert.ToInt32(reader["kol"]);
                                            else count = 0;
                                            if (summa == 0) oplataperiodtemplate.Rows[w].Cells[q + 1].Value = summa + "р; " + count.ToString() + " | ";
                                            else oplataperiodtemplate.Rows[w].Cells[q + 1].Value = summa.ToString("#") + "р; " + count.ToString() + " | ";
                                            totalkol += count;
                                            totalsum += summa;

                                        }
                                    }
                                }
                                
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствТекущееНакопление where Шаблон = '" + data.template()[w] + "' and Подразделение = '" + data.DOanal()[q] + "' and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                                using (var reader = cmd.ExecuteReader())
                                {
                                    if (reader.HasRows)
                                    {
                                        while (reader.Read())
                                        {
                                            if (reader["summa"].ToString() != "") sum = Convert.ToDouble(reader["summa"]);
                                            else sum = 0;
                                            if (reader["kol"].ToString() != "") coun = Convert.ToInt32(reader["kol"]);
                                            else coun = 0;
                                            summa1 = summa - sum;
                                            count1 = count - coun;

                                            if (sum == 0) oplataperiodtemplate.Rows[w].Cells[q + 1].Value += sum + "р; " + coun.ToString() + " | ";
                                            else oplataperiodtemplate.Rows[w].Cells[q + 1].Value += sum.ToString("#") + "р; " + coun.ToString() + " | ";

                                            if (summa1 == 0) oplataperiodtemplate.Rows[w].Cells[q + 1].Value += summa1 + "р;" + count1.ToString();
                                            else oplataperiodtemplate.Rows[w].Cells[q + 1].Value += summa1.ToString("#") + "р;" + count1.ToString();
                                            
                                            totalkol1 += coun;
                                            totalsum1 += sum;
                                            totalkol2 += count1;
                                            totalsum2 += summa1;
                                        }
                                    }
                                }
                               
                            }
                        }
                    }

                    if (totalsum == 0) oplataperiodtemplate.Rows[data.template().Length].Cells[q + 1].Value = totalsum + "р; " + totalkol.ToString() + " | ";
                    else oplataperiodtemplate.Rows[data.template().Length].Cells[q + 1].Value = totalsum.ToString("#") + "р; " + totalkol.ToString() + " | ";

                    if (totalsum1 == 0) oplataperiodtemplate.Rows[data.template().Length].Cells[q + 1].Value += totalsum1 + "р; " + totalkol1.ToString() + " | ";
                    else oplataperiodtemplate.Rows[data.template().Length].Cells[q + 1].Value += totalsum1.ToString("#") + "р; " + totalkol1.ToString() + " | ";

                    if (totalsum2 == 0) oplataperiodtemplate.Rows[data.template().Length].Cells[q + 1].Value += totalsum2 + "р; " + totalkol2.ToString() + " | ";
                    else oplataperiodtemplate.Rows[data.template().Length].Cells[q + 1].Value += totalsum2.ToString("#") + "р; " + totalkol2.ToString() + " | ";

                }


























                
                //for (q = 0; q < data.template().Length; q++)
                //{
                //    using (var conn = new NpgsqlConnection(data.path))
                //    {
                //        conn.Open();
                //        using (var cmd = new NpgsqlCommand())
                //        {
                //            cmd.Connection = conn;
                //            cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствНачало where Шаблон = '" + data.template()[q] + "' and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "'";
                //            using (var reader = cmd.ExecuteReader())
                //            {
                //                if (reader.HasRows)
                //                {
                //                    while (reader.Read())
                //                    {
                //                        if (reader["kol"].ToString() != "")
                //                        {
                //                            oplataplankol[q] = Convert.ToInt32(reader["kol"]);
                //                            count += Convert.ToInt32(reader["kol"]);
                //                        }
                //                        if (reader["summa"].ToString() != "")
                //                        {
                //                            oplataplansum[q] = Convert.ToDouble(reader["summa"]);
                //                            summa += Convert.ToDouble(reader["summa"]);
                //                        }
                //                    }
                //                }
                //            }
                //            cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ВозвратСредствТекущееНакопление where Шаблон = '" + data.template()[q] + "' and ДатаОплаты <= '" + dateend.Value + "' and ДатаОплаты >= '" + datestart.Value + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                //            using (var reader = cmd.ExecuteReader())
                //            {
                //                if (reader.HasRows)
                //                {
                //                    while (reader.Read())
                //                    {
                //                        if (reader["kol"].ToString() != "")
                //                        {
                //                            oplatarealkol[q] = Convert.ToInt32(reader["kol"]);
                //                            count1 += Convert.ToInt32(reader["kol"]);
                //                        }
                //                        if (reader["summa"].ToString() != "")
                //                        {
                //                            oplatarealsum[q] = Convert.ToDouble(reader["summa"]);
                //                            summa1 += Convert.ToDouble(reader["summa"]);
                //                        }
                //                    }
                //                }
                //            }

                //        }
                //    }
                //    coun = oplataplankol[q] - oplatarealkol[q];
                //    sum = oplataplansum[q] - oplatarealsum[q];

                    

                //    if (sum == 0) oplataperiod.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum;
                //    else oplataperiod.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum.ToString("#.#");
                //    if (oplatarealsum[q] == 0) oplataperiod.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q];
                //    else oplataperiod.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q].ToString("#.#");
                //    if (oplataplansum[q] == 0) oplataperiod.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q];
                //    else oplataperiod.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q].ToString("#.#");

                //    double kach = 0, kachsum = 0;
                //    if (oplataplankol[q] != 0)
                //        kach = (double)coun / oplataplankol[q] * 100;
                //    if (coun == 0)
                //        kach = 0;
                //    if (oplataplansum[q] != 0)
                //        kachsum = (double)sum / oplataplansum[q] * 100;
                //    if (sum == 0)
                //        kachsum = 0;
                //    oplataperiod.Rows[q].Cells[4].Value = Math.Round(kach, 0).ToString() + "%  | " + Math.Round(kachsum, 0).ToString() + "%";
                //}
                //oplataperiod.Rows.Add();
                //oplataperiod.Rows[q].Cells[0].Value = "Всего";
                //oplataperiod.Rows[q].Cells[1].Value = count.ToString() + " | " + summa.ToString("#.#");
                //oplataperiod.Rows[q].Cells[2].Value = count1.ToString() + " | " + summa1.ToString("#.#");
                //oplataperiod.Rows[q].Cells[3].Value = (count - count1).ToString() + " | " + (summa - summa1).ToString("#.#");
                //double kach1 = 0, kachsum1 = 0;
                //if (count != 0)
                //    kach1 = (double)(count - count1) / count * 100;
                //if (summa != 0)
                //    kachsum1 = (double)(summa - summa1) / summa * 100;

                //oplataperiod.Rows[q].Cells[4].Value = Math.Round(kach1, 0).ToString() + "%  | " + Math.Round(kachsum1, 0).ToString() + "%";
            }
            
            if (comboBox1.SelectedIndex == 3)
            {
                int count = 0, count1 = 0, coun = 0;
                double summa = 0, summa1 = 0, sum = 0;
                closingperiod.Rows.Clear();
                int[] oplataplankol = new int[data.DOanal().Length];
                double[] oplataplansum = new double[data.DOanal().Length];
                int[] oplatarealkol = new int[data.DOanal().Length];
                double[] oplatarealsum = new double[data.DOanal().Length];
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ЗакрытиеНачало where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаЗакрытия <= '" + dateend.Value + "' and ДатаЗакрытия >= '" + datestart.Value + "'";
                            //}
                            //else 
                            {
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ЗакрытиеНачало where Подразделение = '" + data.DOanal()[q] + "' and ДатаЗакрытия <= '" + dateend.Value + "' and ДатаЗакрытия >= '" + datestart.Value + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplataplankol[q] += Convert.ToInt32(reader["kol"]);
                                            count += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplataplansum[q] += Convert.ToDouble(reader["summa"]);
                                            summa += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ЗакрытиеТекущее where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаЗакрытия <= '" + dateend.Value + "' and ДатаЗакрытия >= '" + datestart.Value + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                            //}
                            //else
                            {
                                cmd.CommandText = "SELECT count(Код) as kol,sum(cast(Сумма as real)) as summa FROM ЗакрытиеТекущее where Подразделение = '" + data.DOanal()[q] + "' and ДатаЗакрытия <= '" + dateend.Value + "' and ДатаЗакрытия >= '" + datestart.Value + "' and ДатаФорм = '" + dateend.Value.Date.AddDays(1) + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplatarealkol[q] += Convert.ToInt32(reader["kol"]);
                                            count1 += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplatarealsum[q] += Convert.ToDouble(reader["summa"]);
                                            summa1 += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }

                        }
                    }
                    coun = oplataplankol[q] - oplatarealkol[q];
                    sum = oplataplansum[q] - oplatarealsum[q];

                    closingperiod.Rows.Add();
                    closingperiod.Rows[q].Cells[0].Value = data.DOanal()[q];

                    if (sum == 0) closingperiod.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum;
                    else closingperiod.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum.ToString("#.#");
                    if (oplatarealsum[q] == 0) closingperiod.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q];
                    else closingperiod.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q].ToString("#.#");
                    if (oplataplansum[q] == 0) closingperiod.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q];
                    else closingperiod.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q].ToString("#.#");

                    double kach = 0, kachsum = 0;
                    if (oplataplankol[q] != 0)
                        kach = (double)coun / oplataplankol[q] * 100;
                    if (coun == 0)
                        kach = 0;
                    if (oplataplansum[q] != 0)
                        kachsum = (double)sum / oplataplansum[q] * 100;
                    if (sum == 0)
                        kachsum = 0;
                    closingperiod.Rows[q].Cells[4].Value = Math.Round(kach, 0).ToString() + "%  | " + Math.Round(kachsum, 0).ToString() + "%";
                }
                closingperiod.Rows.Add();
                closingperiod.Rows[q].Cells[0].Value = "Всего";
                closingperiod.Rows[q].Cells[1].Value = count.ToString() + " | " + summa.ToString("#.#");
                closingperiod.Rows[q].Cells[2].Value = count1.ToString() + " | " + summa1.ToString("#.#");
                closingperiod.Rows[q].Cells[3].Value = (count - count1).ToString() + " | " + (summa - summa1).ToString("#.#");
                double kach1 = 0, kachsum1 = 0;
                if (count != 0)
                    kach1 = (double)(count - count1) / count * 100;
                if (summa != 0)
                    kachsum1 = (double)(summa - summa1) / summa * 100;

                closingperiod.Rows[q].Cells[4].Value = Math.Round(kach1, 0).ToString() + "%  | " + Math.Round(kachsum1, 0).ToString() + "%";
            }
            if (comboBox1.SelectedIndex == 4)
            {
                int count = 0, count1 = 0, coun = 0;
                double summa = 0, summa1 = 0, sum = 0;
                Dolg.Rows.Clear();
                int[] oplataplankol = new int[data.DOanal().Length];
                double[] oplataplansum = new double[data.DOanal().Length];
                int[] oplatarealkol = new int[data.DOanal().Length];
                double[] oplatarealsum = new double[data.DOanal().Length];
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT КолВоДолг as kol ,СуммаДолг as summa FROM АнализСБ where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаФорм = '" + datestart.Value + "'";
                            //}
                            //else 
                            {
                                cmd.CommandText = "SELECT КолВоДолг as kol ,СуммаДолг as summa FROM АнализСБ where Подразделение = '" + data.DOanal()[q] + "' and ДатаФорм = '" + datestart.Value + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplataplankol[q] += Convert.ToInt32(reader["kol"]);
                                            count += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplataplansum[q] += Convert.ToDouble(reader["summa"]);
                                            summa += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT КолВоДолг as kol ,СуммаДолг as summa FROM АнализСБ where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаФорм = '" + dateend.Value + "'";
                            //}
                            //else
                            {
                                cmd.CommandText = "SELECT КолВоДолг as kol ,СуммаДолг as summa FROM АнализСБ where Подразделение = '" + data.DOanal()[q] + "' and ДатаФорм = '" + dateend.Value + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplatarealkol[q] += Convert.ToInt32(reader["kol"]);
                                            count1 += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplatarealsum[q] += Convert.ToDouble(reader["summa"]);
                                            summa1 += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }

                        }
                    }
                    coun = oplataplankol[q] - oplatarealkol[q];
                    sum = oplataplansum[q] - oplatarealsum[q];

                    Dolg.Rows.Add();
                    Dolg.Rows[q].Cells[0].Value = data.DOanal()[q];

                    if (sum == 0) Dolg.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum;
                    else Dolg.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum.ToString("#.#");
                    if (oplatarealsum[q] == 0) Dolg.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q];
                    else Dolg.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q].ToString("#.#");
                    if (oplataplansum[q] == 0) Dolg.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q];
                    else Dolg.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q].ToString("#.#");

                    double kach = 0, kachsum = 0;
                    if (oplataplankol[q] != 0)
                        kach = (double)coun / oplataplankol[q] * 100;
                    if (coun == 0)
                        kach = 0;
                    if (oplataplansum[q] != 0)
                        kachsum = (double)sum / oplataplansum[q] * 100;
                    if (sum == 0)
                        kachsum = 0;
                    Dolg.Rows[q].Cells[4].Value = Math.Round(kach, 0).ToString() + "%  | " + Math.Round(kachsum, 0).ToString() + "%";
                }
                Dolg.Rows.Add();
                Dolg.Rows[q].Cells[0].Value = "Всего";
                Dolg.Rows[q].Cells[1].Value = count.ToString() + " | " + summa.ToString("#.#");
                Dolg.Rows[q].Cells[2].Value = count1.ToString() + " | " + summa1.ToString("#.#");
                Dolg.Rows[q].Cells[3].Value = (count - count1).ToString() + " | " + (summa - summa1).ToString("#.#");
                double kach1 = 0, kachsum1 = 0;
                if (count != 0)
                    kach1 = (double)(count - count1) / count * 100;
                if (summa != 0)
                    kachsum1 = (double)(summa - summa1) / summa * 100;

                Dolg.Rows[q].Cells[4].Value = Math.Round(kach1, 0).ToString() + "%  | " + Math.Round(kachsum1, 0).ToString() + "%";
            }
            if (comboBox1.SelectedIndex == 5)
            {
                int count = 0, count1 = 0, coun = 0;
                double summa = 0, summa1 = 0, sum = 0;
                sevensix.Rows.Clear();
                int[] oplataplankol = new int[data.DOanal().Length];
                double[] oplataplansum = new double[data.DOanal().Length];
                int[] oplatarealkol = new int[data.DOanal().Length];
                double[] oplatarealsum = new double[data.DOanal().Length];
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT КолВоПрет as kol ,СуммаПрет as summa FROM АнализСБ where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаФорм = '" + datestart.Value + "'";
                            //}
                            //else
                            {
                                cmd.CommandText = "SELECT КолВоПрет as kol ,СуммаПрет as summa FROM АнализСБ where Подразделение = '" + data.DOanal()[q] + "' and ДатаФорм = '" + datestart.Value + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplataplankol[q] += Convert.ToInt32(reader["kol"]);
                                            count += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplataplansum[q] += Convert.ToDouble(reader["summa"]);
                                            summa += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }
                            //if (data.DOanal()[q] == "Бирск")
                            //{
                            //    cmd.CommandText = "SELECT КолВоПрет as kol ,СуммаПрет as summa FROM АнализСБ where (Подразделение = '" + data.DOanal()[q] + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаФорм = '" + dateend.Value + "'";
                            //}
                            //else
                            {
                                cmd.CommandText = "SELECT КолВоПрет as kol ,СуммаПрет as summa FROM АнализСБ where Подразделение = '" + data.DOanal()[q] + "' and ДатаФорм = '" + dateend.Value + "'";
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["kol"].ToString() != "")
                                        {
                                            oplatarealkol[q] += Convert.ToInt32(reader["kol"]);
                                            count1 += Convert.ToInt32(reader["kol"]);
                                        }
                                        if (reader["summa"].ToString() != "")
                                        {
                                            oplatarealsum[q] += Convert.ToDouble(reader["summa"]);
                                            summa1 += Convert.ToDouble(reader["summa"]);
                                        }
                                    }
                                }
                            }

                        }
                    }
                    coun = oplataplankol[q] - oplatarealkol[q];
                    sum = oplataplansum[q] - oplatarealsum[q];

                    sevensix.Rows.Add();
                    sevensix.Rows[q].Cells[0].Value = data.DOanal()[q];

                    if (sum == 0) sevensix.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum;
                    else sevensix.Rows[q].Cells[3].Value = coun.ToString() + " | " + sum.ToString("#.#");
                    if (oplatarealsum[q] == 0) sevensix.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q];
                    else sevensix.Rows[q].Cells[2].Value = oplatarealkol[q].ToString() + " | " + oplatarealsum[q].ToString("#.#");
                    if (oplataplansum[q] == 0) sevensix.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q];
                    else sevensix.Rows[q].Cells[1].Value = oplataplankol[q].ToString() + " | " + oplataplansum[q].ToString("#.#");

                    double kach = 0, kachsum = 0;
                    if (oplataplankol[q] != 0)
                        kach = (double)coun / oplataplankol[q] * 100;
                    if (coun == 0)
                        kach = 0;
                    if (oplataplansum[q] != 0)
                        kachsum = (double)sum / oplataplansum[q] * 100;
                    if (sum == 0)
                        kachsum = 0;
                    sevensix.Rows[q].Cells[4].Value = Math.Round(kach, 0).ToString() + "%  | " + Math.Round(kachsum, 0).ToString() + "%";
                }
                sevensix.Rows.Add();
                sevensix.Rows[q].Cells[0].Value = "Всего";
                sevensix.Rows[q].Cells[1].Value = count.ToString() + " | " + summa.ToString("#.#");
                sevensix.Rows[q].Cells[2].Value = count1.ToString() + " | " + summa1.ToString("#.#");
                sevensix.Rows[q].Cells[3].Value = (count - count1).ToString() + " | " + (summa - summa1).ToString("#.#");
                double kach1 = 0, kachsum1 = 0;
                if (count != 0)
                    kach1 = (double)(count - count1) / count * 100;
                if (summa != 0)
                    kachsum1 = (double)(summa - summa1) / summa * 100;

                sevensix.Rows[q].Cells[4].Value = Math.Round(kach1, 0).ToString() + "%  | " + Math.Round(kachsum1, 0).ToString() + "%";
            }
            if (comboBox1.SelectedIndex == 6)
            {
                tablezaemberkontrol.Rows.Clear();
                int[] num = new int[data.users().Length];
                int[] userallmistakes = new int[data.users().Length];
                int[] usercostmistakes = new int[data.users().Length];

                for (q = 0; q < data.users().Length; q++)
                {
                    int relative1 = 0;
                    int i = 0;
                    using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                            cmd.CommandText = "select Дата,ФИОСотрудника,ОтветПроверяющего,ДатаОтветаПроверяющего from ПроверкаМенеджеров where ФИОСотрудника = '" + data.users()[q] + "' and (Дата < '" + dateend.Value.AddDays(1) + "' and Дата >= '" + datestart.Value.AddDays(-1) + "')";
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["ДатаОтветаПроверяющего"] != DBNull.Value)
                                        {
                                            TimeSpan span1 = Convert.ToDateTime(reader["ДатаОтветаПроверяющего"]) - Convert.ToDateTime(reader["Дата"]);
                                            relative1 = span1.Days;
                                        }
                                        else
                                        {
                                            TimeSpan span1 = dateend.Value.AddDays(2) - Convert.ToDateTime(reader["Дата"]);
                                            relative1 = span1.Days;
                                        }
                                        if (relative1 > 0) usercostmistakes[q] += 50;
                                       
                                        usercostmistakes[q] += relative1 * 20;
                                        i++;
                                    }
                                }
                            }
                        }
                    }
                    tablezaemberkontrol.Rows.Add();
                    tablezaemberkontrol.Rows[q].Cells[0].Value = data.users()[q];
                    tablezaemberkontrol.Rows[q].Cells[1].Value =i.ToString();
                    tablezaemberkontrol.Rows[q].Cells[2].Value = usercostmistakes[q].ToString() + " рублей";
                }
                for (int i = 0; i < tablezaemberkontrol.RowCount; i++)
                {
                    if (tablezaemberkontrol[1, i].Value.ToString() == num.Max().ToString())
                    {
                        tablezaemberkontrol.Rows[i].Cells[1].Style.BackColor = Color.Red;
                    }
                    if (tablezaemberkontrol[1, i].Value.ToString() == num.Min().ToString())
                    {
                        tablezaemberkontrol.Rows[i].Cells[1].Style.BackColor = Color.Green;
                    }
                }
            }
            

            if (comboBox1.SelectedIndex == 7)
            {
                tablezaemberkontrol.Rows.Clear();
                int[] num = new int[data.DOanal().Length];
                for (q = 0; q < data.DOanal().Length; q++)
                {
                    Console.WriteLine(data.DOanal()[q]);
                    using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                            cmd.CommandText = "select sum(cast(КоличествоЗаем as int)) as c,sum(cast(СуммаЗаем as int)) as z from Анализсберзаем where Подразделение = '" + data.DOanal()[q] + "' and (Дата between '" + datestart.Value + "' and '" + dateend.Value + "')";
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        tablezaemberkontrol.Rows.Add();
                                        tablezaemberkontrol.Rows[q].Cells[0].Value = data.DOanal()[q];
                                        tablezaemberkontrol.Rows[q].Cells[1].Value = reader["c"].ToString();
                                        tablezaemberkontrol.Rows[q].Cells[2].Value = reader["z"].ToString();
                                        //num[q] = Convert.ToInt32(reader["c"]);
                                    }
                                }
                            }
                        }
                    }
                }
                //for (int i = 0; i < tablezaemberkontrol.RowCount; i++)
                //{
                //    if (tablezaemberkontrol[1, i].Value.ToString() == num.Max().ToString())
                //    {
                //        tablezaemberkontrol.Rows[i].Cells[1].Style.BackColor = Color.Green;
                //    }
                //    if (tablezaemberkontrol[1, i].Value.ToString() == num.Min().ToString())
                //    {
                //        tablezaemberkontrol.Rows[i].Cells[1].Style.BackColor = Color.Red;
                //    }
                //}
            }
            if (comboBox1.SelectedIndex == 8)
            {
                string nocall = "0";
                string norecord = "0";
                tablecalls.Rows.Clear();
                int[] num = new int[data.DOanal().Length];
                for (q = 0; q < data.DOanal().Length; q++)
                {
                     using (var conn = new NpgsqlConnection(data.path))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand())
                        {
                            cmd.Connection = conn;
                           
                            cmd.CommandText = "select count(ФИОЗаемщика) as c from Оплата where Оплата.Договор not in (select Договор from РаботаДолг where Оплата.Договор = РаботаДолг.Договор and Оплата.ФИОЗаемщика = РаботаДолг.ФИО and ДатаРаботы>='" + datestart.Value.Date + "' and ДатаРаботы<='" + datestart.Value.Date.AddDays(1).AddTicks(-1) + "')and (cast(ДнейПросрочки as int) > '0' or ДатаВозврата = '" + datestart.Value.Date + "') and Подразделение = '" + data.DOanal()[q] + "'";
                           
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        nocall = reader["c"].ToString();
                                       
                                    }
                                }
                            }
                            cmd.CommandText = "select count(ФИОЗаемщика) as c from Оплата where Оплата.Договор in (select Договор from РаботаДолг where Оплата.Договор = РаботаДолг.Договор and Оплата.ФИОЗаемщика = РаботаДолг.ФИО and ДатаРаботы>'=" + datestart.Value.Date + "' and ДатаРаботы<='" + datestart.Value.Date.AddDays(1).AddTicks(-1) + "' and Запись='')and (cast(ДнейПросрочки as int) > '0' or ДатаВозврата = '" + datestart.Value.Date + "') and Подразделение = '" + data.DOanal()[q] + "'";

                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        norecord = reader["c"].ToString();
                                        
                                    }
                                }
                            }
                        }
                    }
                     tablecalls.Rows.Add();
                     tablecalls.Rows[q].Cells[0].Value = data.DOanal()[q];
                     tablecalls.Rows[q].Cells[1].Value = nocall;
                     tablecalls.Rows[q].Cells[2].Value = norecord;
                }

               
            }
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dateend.Visible = true;
            datestart.Visible = true;
            var yr = DateTime.Today.Year;
            var mth = DateTime.Today.Month;
            var firstDay = new DateTime(yr, mth, 1);
            var lastDay = new DateTime(yr, mth, 1).AddMonths(1).AddDays(-1);
            datestart.Value = firstDay;
            dateform.Value = DateTime.Today;
            if (firstDay == DateTime.Today)
            {
                dateend.Value = DateTime.Today.AddDays(1);
            }
            else
            {
                dateend.Value = DateTime.Today.AddDays(-1);
            }

            if (comboBox1.SelectedIndex == 0)
            {
                visible(oplatadays);
            }
             if (comboBox1.SelectedIndex == 2)
            {
                visible(oplataperiodtemplate);
            }
            
            if (comboBox1.SelectedIndex == 1 )
            {
                visible(oplataperiod);
            }
            if (comboBox1.SelectedIndex == 3)
            {
                visible(closingperiod);
            }
            if (comboBox1.SelectedIndex == 4)
            {
                visible(Dolg);
            }
            if (comboBox1.SelectedIndex ==5)
            {
                visible(sevensix);
            }
            if (comboBox1.SelectedIndex == 6 || comboBox1.SelectedIndex == 7)
            {
                visible(tablezaemberkontrol);
            }
            if (comboBox1.SelectedIndex == 8)
            {
                visible(tablecalls);
            }
        }

        private void visible(DataGridView data)
        {
            foreach (Control control in this.Controls)
                if (control is DataGridView)
                {
                    control.Visible = false;
                }
            data.Visible = true;
        }

        private void dateend_ValueChanged(object sender, EventArgs e)
        {
            dateform.Value = dateend.Value.AddDays(1);
        }

        
  
        private void closingperiod_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            data.whatanaliz = "По периодам закрытие";
            data.analizpodr = closingperiod.Rows[closingperiod.CurrentRow.Index].Cells[0].Value.ToString();
            data.analizdatestart = datestart.Value;
            data.analizdateend = dateend.Value;
            data.analizdateform = dateend.Value.Date.AddDays(1);
            new analizprint().ShowDialog();
        }

        private void oplatadays_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            data.whatanaliz = "По дням";
            data.analizpodr = oplatadays.Columns[oplatadays.CurrentCell.ColumnIndex].Name.ToString();
            data.analizdateotch = Convert.ToDateTime(oplatadays.Rows[oplatadays.CurrentRow.Index].Cells[0].Value).ToString();
            data.analizdateform = dateend.Value.Date.AddDays(1);
            new analizprint().ShowDialog();
        }

        private void oplataperiod_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            data.whatanaliz = "По периодам";
            data.analizpodr = oplataperiod.Rows[oplataperiod.CurrentRow.Index].Cells[0].Value.ToString();
            data.analizdatestart = datestart.Value;
            data.analizdateend = dateend.Value;
            data.analizdateform = dateend.Value.Date.AddDays(1);
            new analizprint().ShowDialog();
        }



        private void oplataperiod_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.ClipBounds, true);
            Font font = e.CellStyle.Font;
            using (var sf = new StringFormat())
            using (var firstBrush = new SolidBrush(e.CellStyle.ForeColor))
            using (var lastBrush = new SolidBrush(Color.Red))
            using (var boldFont = new Font(font, FontStyle.Bold))
            {
                sf.LineAlignment = StringAlignment.Center;
                sf.FormatFlags = StringFormatFlags.NoWrap;
                sf.Trimming = StringTrimming.EllipsisWord;

                string text = (string)e.FormattedValue;



                if (text != null)
                {
                    if (text.Contains(" | ") && e.RowIndex >= 0)
                    {
                        String[] words = text.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string lastHalf = " | ";
                        string str = "";
                        var rect1 = new RectangleF();
                        for (int i = 0; i != words.Length; i++)
                        {
                            SizeF size = e.Graphics.MeasureString("11111", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                            var rect = new RectangleF(e.CellBounds.X + size.Width, e.CellBounds.Y, e.CellBounds.Width - size.Width, e.CellBounds.Height);
                            if (i == 0)
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, e.CellBounds.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(e.CellBounds.X + size1.Width, e.CellBounds.Y, e.CellBounds.Width - size1.Width, e.CellBounds.Height);
                            }
                            else
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, rect.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(rect.X + 20, e.CellBounds.Y, rect.Width - 20, e.CellBounds.Height);
                            }

                            e.Graphics.DrawString(words[i], font, firstBrush, rect1, sf);

                            if (words.Length - i != 1)
                            {
                                e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect, sf);
                                str += words[i] + lastHalf;
                            }
                        }
                    }
                    else
                    {
                        e.Graphics.DrawString(text, font, firstBrush, e.CellBounds, sf);
                    }
                }
            }
            e.Handled = true;
        }

        private void oplatadays_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.ClipBounds, true);
            Font font = e.CellStyle.Font;
            using (var sf = new StringFormat())
            using (var firstBrush = new SolidBrush(e.CellStyle.ForeColor))
            using (var lastBrush = new SolidBrush(Color.Red))
            using (var boldFont = new Font(font, FontStyle.Bold))
            {
                sf.LineAlignment = StringAlignment.Center;

                string text = (string)e.FormattedValue;



                if (text != null)
                {
                    if (text.Contains(" | ") && e.RowIndex >= 0)
                    {
                        String[] words = text.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string lastHalf = " | ";
                        RectangleF rect, rect1, rect2, rect3, rect4 = new RectangleF();

                        SizeF size = e.Graphics.MeasureString("", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect = new RectangleF(e.CellBounds.X + size.Width, e.CellBounds.Y, e.CellBounds.Width - size.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(words[0], font, firstBrush, rect, sf);

                        SizeF size1 = e.Graphics.MeasureString(words[0], font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect1 = new RectangleF(e.CellBounds.X + size1.Width, e.CellBounds.Y, e.CellBounds.Width - size1.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect1, sf);

                        SizeF size2 = e.Graphics.MeasureString("1", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect2 = new RectangleF(rect1.X + size2.Width, e.CellBounds.Y, rect1.Width - size2.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(words[1], font, firstBrush, rect2, sf);

                        SizeF size3 = e.Graphics.MeasureString(words[1], font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect3 = new RectangleF(rect2.X + size3.Width, e.CellBounds.Y, rect2.Width - size3.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect3, sf);

                        SizeF size4 = e.Graphics.MeasureString("1", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect4 = new RectangleF(rect3.X + size4.Width, e.CellBounds.Y, rect3.Width - size4.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(words[2], font, firstBrush, rect4, sf);

                    }
                    else
                    {
                        e.Graphics.DrawString(text, font, firstBrush, e.CellBounds, sf);
                    }
                }
            }
            e.Handled = true;
        }

        private void closingperiod_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.ClipBounds, true);
            Font font = e.CellStyle.Font;
            using (var sf = new StringFormat())
            using (var firstBrush = new SolidBrush(e.CellStyle.ForeColor))
            using (var lastBrush = new SolidBrush(Color.Red))
            using (var boldFont = new Font(font, FontStyle.Bold))
            {
                sf.LineAlignment = StringAlignment.Center;
                sf.FormatFlags = StringFormatFlags.NoWrap;
                sf.Trimming = StringTrimming.EllipsisWord;

                string text = (string)e.FormattedValue;



                if (text != null)
                {
                    if (text.Contains(" | ") && e.RowIndex >= 0)
                    {
                        String[] words = text.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string lastHalf = " | ";
                        string str = "";
                        var rect1 = new RectangleF();
                        for (int i = 0; i != words.Length; i++)
                        {
                            SizeF size = e.Graphics.MeasureString("11111", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                            var rect = new RectangleF(e.CellBounds.X + size.Width, e.CellBounds.Y, e.CellBounds.Width - size.Width, e.CellBounds.Height);
                            if (i == 0)
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, e.CellBounds.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(e.CellBounds.X + size1.Width, e.CellBounds.Y, e.CellBounds.Width - size1.Width, e.CellBounds.Height);
                            }
                            else
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, rect.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(rect.X + 20, e.CellBounds.Y, rect.Width - 20, e.CellBounds.Height);
                            }

                            e.Graphics.DrawString(words[i], font, firstBrush, rect1, sf);

                            if (words.Length - i != 1)
                            {
                                e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect, sf);
                                str += words[i] + lastHalf;
                            }
                        }
                    }
                    else
                    {
                        e.Graphics.DrawString(text, font, firstBrush, e.CellBounds, sf);
                    }
                }
            }
            e.Handled = true;
        }

        private void Dolg_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.ClipBounds, true);
            Font font = e.CellStyle.Font;
            using (var sf = new StringFormat())
            using (var firstBrush = new SolidBrush(e.CellStyle.ForeColor))
            using (var lastBrush = new SolidBrush(Color.Red))
            using (var boldFont = new Font(font, FontStyle.Bold))
            {
                sf.LineAlignment = StringAlignment.Center;
                sf.FormatFlags = StringFormatFlags.NoWrap;
                sf.Trimming = StringTrimming.EllipsisWord;

                string text = (string)e.FormattedValue;



                if (text != null)
                {
                    if (text.Contains(" | ") && e.RowIndex >= 0)
                    {
                        String[] words = text.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string lastHalf = " | ";
                        string str = "";
                        var rect1 = new RectangleF();
                        for (int i = 0; i != words.Length; i++)
                        {
                            SizeF size = e.Graphics.MeasureString("11111", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                            var rect = new RectangleF(e.CellBounds.X + size.Width, e.CellBounds.Y, e.CellBounds.Width - size.Width, e.CellBounds.Height);
                            if (i == 0)
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, e.CellBounds.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(e.CellBounds.X + size1.Width, e.CellBounds.Y, e.CellBounds.Width - size1.Width, e.CellBounds.Height);
                            }
                            else
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, rect.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(rect.X + 20, e.CellBounds.Y, rect.Width - 20, e.CellBounds.Height);
                            }

                            e.Graphics.DrawString(words[i], font, firstBrush, rect1, sf);

                            if (words.Length - i != 1)
                            {
                                e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect, sf);
                                str += words[i] + lastHalf;
                            }
                        }
                    }
                    else
                    {
                        e.Graphics.DrawString(text, font, firstBrush, e.CellBounds, sf);
                    }
                }
            }
            e.Handled = true;
        }

        private void sevensix_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.ClipBounds, true);
            Font font = e.CellStyle.Font;
            using (var sf = new StringFormat())
            using (var firstBrush = new SolidBrush(e.CellStyle.ForeColor))
            using (var lastBrush = new SolidBrush(Color.Red))
            using (var boldFont = new Font(font, FontStyle.Bold))
            {
                sf.LineAlignment = StringAlignment.Center;
                sf.FormatFlags = StringFormatFlags.NoWrap;
                sf.Trimming = StringTrimming.EllipsisWord;

                string text = (string)e.FormattedValue;



                if (text != null)
                {
                    if (text.Contains(" | ") && e.RowIndex >= 0)
                    {
                        String[] words = text.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string lastHalf = " | ";
                        string str = "";
                        var rect1 = new RectangleF();
                        for (int i = 0; i != words.Length; i++)
                        {
                            SizeF size = e.Graphics.MeasureString("11111", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                            var rect = new RectangleF(e.CellBounds.X + size.Width, e.CellBounds.Y, e.CellBounds.Width - size.Width, e.CellBounds.Height);
                            if (i == 0)
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, e.CellBounds.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(e.CellBounds.X + size1.Width, e.CellBounds.Y, e.CellBounds.Width - size1.Width, e.CellBounds.Height);
                            }
                            else
                            {
                                SizeF size1 = e.Graphics.MeasureString(str, font, rect.Location, StringFormat.GenericTypographic);
                                rect1 = new RectangleF(rect.X + 20, e.CellBounds.Y, rect.Width - 20, e.CellBounds.Height);
                            }

                            e.Graphics.DrawString(words[i], font, firstBrush, rect1, sf);

                            if (words.Length - i != 1)
                            {
                                e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect, sf);
                                str += words[i] + lastHalf;
                            }
                        }
                    }
                    else
                    {
                        e.Graphics.DrawString(text, font, firstBrush, e.CellBounds, sf);
                    }
                }
            }
            e.Handled = true;
        }

        private void tablecalls_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void oplataperiodtemplate_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.ClipBounds, true);
            Font font = e.CellStyle.Font;
            using (var sf = new StringFormat())
            using (var firstBrush = new SolidBrush(e.CellStyle.ForeColor))
            using (var lastBrush = new SolidBrush(Color.Red))
            using (var boldFont = new Font(font, FontStyle.Bold))
            {
                sf.LineAlignment = StringAlignment.Center;

                string text = (string)e.FormattedValue;



                if (text != null)
                {
                    if (text.Contains(" | ") && e.RowIndex >= 0)
                    {
                        String[] words = text.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string lastHalf = " | ";
                        RectangleF rect, rect1, rect2, rect3, rect4 = new RectangleF();

                        SizeF size = e.Graphics.MeasureString("", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect = new RectangleF(e.CellBounds.X + size.Width, e.CellBounds.Y, e.CellBounds.Width - size.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(words[0], font, firstBrush, rect, sf);

                        SizeF size1 = e.Graphics.MeasureString(words[0], font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect1 = new RectangleF(e.CellBounds.X + size1.Width, e.CellBounds.Y, e.CellBounds.Width - size1.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect1, sf);

                        SizeF size2 = e.Graphics.MeasureString("1", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect2 = new RectangleF(rect1.X + size2.Width, e.CellBounds.Y, rect1.Width - size2.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(words[1], font, firstBrush, rect2, sf);

                        SizeF size3 = e.Graphics.MeasureString(words[1], font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect3 = new RectangleF(rect2.X + size3.Width, e.CellBounds.Y, rect2.Width - size3.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(lastHalf, boldFont, lastBrush, rect3, sf);

                        SizeF size4 = e.Graphics.MeasureString("1", font, e.CellBounds.Location, StringFormat.GenericTypographic);
                        rect4 = new RectangleF(rect3.X + size4.Width, e.CellBounds.Y, rect3.Width - size4.Width, e.CellBounds.Height);
                        e.Graphics.DrawString(words[2], font, firstBrush, rect4, sf);

                    }
                    else
                    {
                        e.Graphics.DrawString(text, font, firstBrush, e.CellBounds, sf);
                    }
                }
            }
            e.Handled = true;
        }

        private void tablecalls_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            data.whatanaliz = "Звонки";
            data.analizpodr = tablecalls.Rows[tablecalls.CurrentRow.Index].Cells[0].Value.ToString();
            data.analizdateend = datestart.Value;
            new analizprint().ShowDialog();
        }




    }

}
