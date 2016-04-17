using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CountMoney
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //Thread to show splash window
            Thread thUI = new Thread(new ThreadStart(ShowSplashWindow));
            thUI.Name = "Splash UI";
            thUI.Priority = ThreadPriority.Normal;
            thUI.IsBackground = true;
            thUI.Start();

            //Thread to load time-consuming resources.
            Thread th = new Thread(new ThreadStart(LoadResources));
            th.Name = "Resource Loader";
            th.Priority = ThreadPriority.Highest;
            th.Start();

            th.Join();

            if (SplashForm != null)
            {
                SplashForm.Invoke(new MethodInvoker(delegate { SplashForm.Close(); }));
            }

            thUI.Join();

            Application.Run(new MainForm());
        }
        public static WelcomeForm SplashForm
        {
            get;
            set;
        }
        
        private static void LoadResources()
        {
            for (int i = 1; i <= 15; i++)
            {
                if (SplashForm != null)
                {
                    SplashForm.Invoke(new MethodInvoker(delegate { SplashForm.status.Text = Properties.Resources.String1; }));
                }
                Thread.Sleep(100);
            }
            SplashForm.Invoke(new MethodInvoker(delegate { SplashForm.status.Text = "完成" + DateTime.Now.ToString(); }));
        }
        private static void ShowSplashWindow()
        {
            SplashForm = new WelcomeForm();
            Application.Run(SplashForm);
        }
    }
}
