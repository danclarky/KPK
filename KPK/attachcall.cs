using MySql.Data.MySqlClient;
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
    public partial class attachcall : Form
    {
        public attachcall()
        {
            InitializeComponent(); ;
        }

        private void attachcall_Load(object sender, EventArgs e)
        {
            listofcalls.Rows.Clear();
            using (var conn = new MySqlConnection(data.stringconnect()[0]))
            {
                conn.Open();
                using (var cmd = new MySqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "SELECT src,dst ,disposition ,calldate,billsec,recordingfile FROM cdr where src = '" + data.userphone + "' order by calldate desc";
                    using (var reader = cmd.ExecuteReader())
                    {
                        int i = 0;
                        while (reader.Read())
                        {
                            listofcalls.Rows.Add();
                            listofcalls.Rows[i].Cells[0].Value = reader["dst"].ToString();
                            listofcalls.Rows[i].Cells[1].Value = Convert.ToDateTime(reader["calldate"]); 
                            listofcalls.Rows[i].Cells[2].Value = reader["billsec"].ToString();
                            listofcalls.Rows[i].Cells[3].Value = reader["disposition"].ToString();
                            listofcalls.Rows[i].Cells[4].Value = reader["recordingfile"].ToString();
                            i++;
                        }
                    }
                }
            }
        }

        private void listofcalls_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string txt = listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[1].Value.ToString();
                data.callaudio = "\\\\78.69.157.1\\calls\\" + txt.Substring(6, 4) + "\\" + txt.Substring(3, 2) + "\\" + txt.Substring(0, 2) + "\\";
                data.callaudio += listofcalls.Rows[listofcalls.CurrentRow.Index].Cells[4].Value.ToString();

                if (data.callaudio != "")
                {
               
                }
                this.Close();
        }
    }
}
