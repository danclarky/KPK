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
    public partial class usersspy : Form
    {
        public usersspy()
        {
            InitializeComponent();
        }

        private void usersspy_Load(object sender, EventArgs e)
        {
            string[,] operations;
            this.Text = "Операции по " + data.usercode;
            dataGridView1.Rows.Clear();
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "Select Время, Действие From ОтслежВходВыход Where Пользователь = '" + data.usercode + "' order by Время desc";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                         i = 0;
                        while (reader.Read())
                        {
                            i++;
                        }
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        operations = new string[i, 2];
                        int q = 0;
                        while (reader.Read())
                        {
                            operations[q, 0] = reader["Время"].ToString();
                            operations[q, 1] = reader["Действие"].ToString();
                            q++;
                        }
                    }
                }
            }
            double raznica = 0;
            for (int i = 0; i < operations.GetLength(0); i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = operations[i, 0];
                dataGridView1.Rows[i].Cells[1].Value = operations[i, 1];
                if (operations[i, 1] == "Выход" && operations[i+1, 1] == "Вход" )
                {
                    TimeSpan  razn = Convert.ToDateTime(operations[i, 0]) - Convert.ToDateTime(operations[i + 1, 0]);
                    raznica += razn.TotalSeconds;
                }

            }
            MessageBox.Show(ToDateTimeDiff(raznica));
        }
        public string ToDateTimeDiff(double Day)
        {
            int hour = Convert.ToInt32(Day) / 3600;
            double minute = (Day % 3600) / 60;
            double second = Day % 60;
            return string.Format("{0} часов {1} минут {2} секунд", hour, Convert.ToInt32(minute), Convert.ToInt32(second));
        }
    }
}
