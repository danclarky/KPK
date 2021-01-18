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

namespace KPK
{
    public partial class analizprint : Form
    {
        public analizprint()
        {
            InitializeComponent();
        }

        private void analizprint_Load(object sender, EventArgs e)
        {
            if (data.whatanaliz == "Звонки")
            {
                oplatavivod.Rows.Clear();
                oplatavivod.Columns.Clear();
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "БезЗвонка", HeaderText = "Без Звонка", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Беззаписи", HeaderText = "Без записи", Width = 200 });
            
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        int q = 0;
                        cmd.Connection = conn;
                        cmd.CommandText = "select ФИОЗаемщика from Оплата where Оплата.Договор not in (select Договор from РаботаДолг where Оплата.Договор = РаботаДолг.Договор and Оплата.ФИОЗаемщика = РаботаДолг.ФИО and ДатаРаботы>='" + data.analizdateend.Date + "' and ДатаРаботы<='" + data.analizdateend.Date.AddDays(1).AddTicks(-1) + "')and (cast(ДнейПросрочки as int) > '0' or ДатаВозврата = '" + data.analizdateend.Date + "') and Подразделение = '" + data.analizpodr + "'";

                        Console.WriteLine(cmd.CommandText);    
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                
                                while (reader.Read())
                                {
                                    oplatavivod.Rows.Add();
                                    oplatavivod.Rows[q].Cells[0].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        cmd.CommandText = "select ФИОЗаемщика from Оплата where Оплата.Договор in (select Договор from РаботаДолг where Оплата.Договор = РаботаДолг.Договор and Оплата.ФИОЗаемщика = РаботаДолг.ФИО and ДатаРаботы>'=" + data.analizdateend.Date + "' and ДатаРаботы<='" + data.analizdateend.Date.AddDays(1).AddTicks(-1) + "' and Запись='')and (cast(ДнейПросрочки as int) > '0' or ДатаВозврата = '" + data.analizdateend.Date + "') and Подразделение = '" + data.analizpodr + "'";

                        Console.WriteLine(cmd.CommandText);
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int w = 0;
                                while (reader.Read())
                                {
                                    if (w >= q)
                                    {
                                        oplatavivod.Rows.Add();
                                    }
                                    oplatavivod.Rows[w].Cells[1].Value = reader["ФИОЗаемщика"].ToString();
                                    w++;
                                }
                            }
                        }
                    }
                }
            }
            if (data.whatanaliz == "По дням")
            {
                oplatavivod.Rows.Clear();
                oplatavivod.Columns.Clear();
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "План", HeaderText = "План", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Выпало", HeaderText = "Выпало", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Оплатило", HeaderText = "Оплатило", Width = 200 });
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT ФИОЗаемщика FROM ВозвратСредствНачало where Подразделение = '" + data.analizpodr + "' and ДатаОплаты = '" + data.analizdateotch + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows.Add();
                                    oplatavivod.Rows[q].Cells[0].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        cmd.CommandText = "SELECT ФИОЗаемщика FROM ВозвратСредствТекущееНакопление where Подразделение = '" + data.analizpodr + "' and ДатаОплаты = '" + data.analizdateotch + "' and ДатаФорм = '" + data.analizdateform + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows[q].Cells[1].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        cmd.CommandText = "select ФИОЗаемщика from ВозвратСредствНачало where ФИОЗаемщика not in (select ФИОЗаемщика from ВозвратСредствТекущееНакопление where ДатаОплаты = '" + data.analizdateotch + "' and ДатаФорм = '" + data.analizdateform + "')  and Подразделение = '" + data.analizpodr + "' and ДатаОплаты = '" + data.analizdateotch + "'";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows[q].Cells[2].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        
                    }
                }
            }
            if (data.whatanaliz == "По периодам")
            {
                oplatavivod.Rows.Clear();
                oplatavivod.Columns.Clear();
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "План", HeaderText = "План", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Выпало", HeaderText = "Выпало", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Оплатило", HeaderText = "Оплатило", Width = 200 });
                
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        if (data.analizpodr == "Бирск")
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ВозвратСредствНачало where (Подразделение = '" + data.analizpodr + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "'";
                        }
                        else
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ВозвратСредствНачало where Подразделение = '" + data.analizpodr + "' and ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "'";
                        }


                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows.Add();
                                    oplatavivod.Rows[q].Cells[0].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        

                        if (data.analizpodr == "Бирск")
                        {
                            
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ВозвратСредствТекущееНакопление where (Подразделение = '" + data.analizpodr + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "'";
                        }
                        else
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ВозвратСредствТекущееНакопление where Подразделение = '" + data.analizpodr + "' and ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "'";
                        }

                        Console.WriteLine(cmd.CommandText);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows[q].Cells[1].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }

                        if (data.analizpodr == "Бирск")
                        {

                         
                            cmd.CommandText = "select ФИОЗаемщика from ВозвратСредствНачало where ФИОЗаемщика not in (select ФИОЗаемщика from ВозвратСредствТекущееНакопление where ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "')  and (Подразделение = '" + data.analizpodr + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "'";
                        }
                        else
                        {
                            cmd.CommandText = "select ФИОЗаемщика from ВозвратСредствНачало where ФИОЗаемщика not in (select ФИОЗаемщика from ВозвратСредствТекущееНакопление where ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "')  and Подразделение = '" + data.analizpodr + "' and ДатаОплаты <= '" + data.analizdateend + "' and ДатаОплаты >= '" + data.analizdatestart + "'";
                        }

                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows[q].Cells[2].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }

                    }
                }
            }
            if (data.whatanaliz == "По периодам закрытие")
            {
                oplatavivod.Rows.Clear();
                oplatavivod.Columns.Clear();
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "План", HeaderText = "План", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Выпало", HeaderText = "Выпало", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Оплатило", HeaderText = "Оплатило", Width = 200 });

                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        if (data.analizpodr == "Бирск")
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ЗакрытиеНачало where (Подразделение = '" + data.analizpodr + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "'";
                        }
                        else
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ЗакрытиеНачало where Подразделение = '" + data.analizpodr + "' and ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "'";
                        }
                       
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows.Add();
                                    oplatavivod.Rows[q].Cells[0].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        if (data.analizpodr == "Бирск")
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ЗакрытиеТекущее where (Подразделение = '" + data.analizpodr + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "'";
                        }
                        else
                        {
                            cmd.CommandText = "SELECT ФИОЗаемщика FROM ЗакрытиеТекущее where Подразделение = '" + data.analizpodr + "' and ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "'";
                        }
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows[q].Cells[1].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }
                        if (data.analizpodr == "Бирск")
                        {
                            cmd.CommandText = "select ФИОЗаемщика from ЗакрытиеНачало where ФИОЗаемщика not in (select ФИОЗаемщика from ЗакрытиеТекущее where ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "')  and (Подразделение = '" + data.analizpodr + "' or Подразделение = 'Бирск ДО №13' or Подразделение = 'Основное подразделение') and ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "'";
                        }
                        else
                        {
                            cmd.CommandText = "select ФИОЗаемщика from ЗакрытиеНачало where ФИОЗаемщика not in (select ФИОЗаемщика from ЗакрытиеТекущее where ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "' and ДатаФорм = '" + data.analizdateform + "')  and Подразделение = '" + data.analizpodr + "' and ДатаЗакрытия <= '" + data.analizdateend + "' and ДатаЗакрытия >= '" + data.analizdatestart + "'";
                        }

                    
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    oplatavivod.Rows[q].Cells[2].Value = reader["ФИОЗаемщика"].ToString();
                                    q++;
                                }
                            }
                        }

                    }
                }
            }
            if (data.whatanaliz == "больше шестид")
            {
                oplatavivod.Rows.Clear();
                oplatavivod.Columns.Clear();
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "План1", HeaderText = "#", Width = 50 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "План", HeaderText = "ФИО", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Выпало", HeaderText = "Подразделение", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Вып1ало", HeaderText = "ДатаПлатежа", Width = 200 });
                oplatavivod.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Сумма", HeaderText = "Сумма", Width = 200 });
                using (var conn = new NpgsqlConnection(data.path))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;
                        //cmd.CommandText = "select ФИОЗаемщика, Подразделение,Договор from Должники where ДнейПросрочки > '60' and ФИОЗаемщика not in (select ФИО from РаботаДолг where Тип ='претензия') order by Подразделение";
                         //cmd.CommandText = "select ФИОЗаемщика, Подразделение,Договор,ДатаПлатежа from семьшесть where Долг < '2017-04-23 00:00:00' order by Подразделение";
                       cmd.CommandText = "select ФИОЗаемщика,ДатаПлатежа,ДолгСуд,Договор, Подразделение from семьшесть where ФИОЗаемщика in (select ФИО from РаботаДолг where position('уме' In Результат) <> '0') order by Подразделение";
                        // cmd.CommandText = "select ФИОЗаемщика,ДатаПослПлатежа,ДолгПоСумме, Подразделение,Договор from Должники where ФИОЗаемщика in (select ФИО from РаботаДолг where position('уме' In Результат) <> '0') order by Подразделение";
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int q = 0;
                                while (reader.Read())
                                {
                                    var j = 0;
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
                                    if (j == 0)
                                    {
                                        oplatavivod.Rows.Add();
                                        oplatavivod.Rows[q].Cells[0].Value = q+1;
                                        oplatavivod.Rows[q].Cells[1].Value = reader["ФИОЗаемщика"].ToString();
                                        oplatavivod.Rows[q].Cells[2].Value = reader["Подразделение"].ToString();
                                        oplatavivod.Rows[q].Cells[3].Value = reader["ДатаПлатежа"].ToString();
                                        oplatavivod.Rows[q].Cells[4].Value = reader["ДолгСуд"].ToString();
                                        q++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
           
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sb = new StringBuilder("<html><head><meta charset='utf-8'><style>table {border-collapse: collapse;}th {border: 1px solid black; padding: 2px;} td {border: 1px solid black;padding: 2px;} </style>")
            .Append("</head><body><table><tr>");
            foreach (DataGridViewColumn c in oplatavivod.Columns)
                sb.Append("<th>").Append(c.HeaderText).Append("</th>");
            foreach (DataGridViewRow o in oplatavivod.Rows)
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
            var pdfgen = new NReco.PdfGenerator.HtmlToPdfConverter();
            pdfgen.GeneratePdf(sb.ToString(), null, "index.pdf");

        }
    }
}
