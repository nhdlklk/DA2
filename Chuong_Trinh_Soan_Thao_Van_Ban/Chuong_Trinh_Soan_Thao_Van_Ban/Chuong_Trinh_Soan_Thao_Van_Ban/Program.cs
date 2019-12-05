using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Chuong_Trinh_Soan_Thao_Van_Ban
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmworkpad());
        }
    }
}
