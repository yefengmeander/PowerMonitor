using System;
using System.Linq;
using System.Windows.Forms;
using log4net;
using log4net.Config;
using System.IO;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
//[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config", Watch = true)]
namespace PowerMonitor
{
    static class Program
    {
        private static ILog _logger = LogManager.GetLogger(typeof(Program));
        private static System.Threading.Mutex mutex;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            mutex = new System.Threading.Mutex(true, "OnlyRun");
            if (mutex.WaitOne(0, false))
            {
            	//InitLog4Net();
                _logger.Info("监控程序启动中...");              
                Application.Run(new Form1());
                _logger.Info("监控程序退出。");
            }
            else
            {
                MessageBox.Show("程序已经在运行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }
        }
        
        /*未使用*/
        private static void InitLog4Net()
        {
            XmlConfigurator.ConfigureAndWatch(
                new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log4net.config")));
        }

    }
}
