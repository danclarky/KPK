using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ServiceProcess;
using System.Management;
using System.Text.RegularExpressions;

namespace KPK
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {

           

            if (System.Diagnostics.Process.GetProcessesByName(Application.ProductName).Length > 1)
            {
                MessageBox.Show("Приложение уже запущено");
                Application.Exit();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                string c = null;
                c = new StreamReader("C:\\КПК\\Программа\\path.ini").ReadLine();
                if (c == "local")
                {
                    data.path = "***";
                    data.ipfiles = "***";
                }
                else if (c == "vpn")
                {
                    data.path = "**";
                    data.ipfiles = "***";
                }
                Application.Run(new mainform());
            }
        }
    }
}
