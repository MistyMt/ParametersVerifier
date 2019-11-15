using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Interface.Forms;


namespace Interface
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new banner1());
            Application.Run(new banner2());
            Application.Run(new ObjectSelect());
            Application.Run(new GeneralInfo());
            Application.Run(new MainForm());
        }
    }
}
