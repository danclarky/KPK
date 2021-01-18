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
    public partial class attach : Form
    {
        public attach()
        {
            InitializeComponent();
        }

        private void attach_Load(object sender, EventArgs e)
        {
            String[] words = data.options.attachfile.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            String[] words1 = data.options.namefile.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i != words.Length; i++)
            {

                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = words1[i];
                dataGridView1.Rows[i].Cells[1].Value = words[i];
            }
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string cellid = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();

           
            Console.WriteLine(cellid);
            System.Diagnostics.Process.Start(cellid);
        }

        private void attach_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
            }
        }
    }
}
