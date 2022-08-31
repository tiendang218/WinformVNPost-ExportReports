using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace XuatExcelApp
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            //MySolutionWinApplication winApplication = new MySolutionWinApplication();
            ////...
            //if (ConfigurationManager.ConnectionStrings["ConnectionString"] != null)
            //{
            //    winApplication.ConnectionString = ConfigurationManager.
            //       ConnectionStrings["ConnectionString"].ConnectionString;
            //}
            //winApplication.Setup();
            //winApplication.Start();
            ////... = new MySolutionWinApplication();
            ////...
            //if (ConfigurationManager.ConnectionStrings["ConnectionString"] != null)
            //{
            //    winApplication.ConnectionString = ConfigurationManager.
            //       ConnectionStrings["ConnectionString"].ConnectionString;
            //}
            //winApplication.Setup();
            //winApplication.Start();
            ////...
        }
    }
}
