using MosTrainer.Services;
using MosTrainer.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MosTrainer
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.ProcessExit += (s, e) => AssetsDeployer.CleanupAll();
            Application.ApplicationExit += (s, e) => AssetsDeployer.CleanupAll();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
          //  Application.Run(new Form1());
            Application.Run(new LoginForm());
        }
    }
}
