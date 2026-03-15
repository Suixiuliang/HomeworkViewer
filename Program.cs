using System;
using System.Threading;
using System.Windows.Forms;

namespace HomeworkViewer
{
    static class Program
    {
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            const string appName = "HomeworkViewer";
            bool createdNew;

            mutex = new Mutex(true, appName, out createdNew);
            if (!createdNew)
            {
                MessageBox.Show("作业展板已经在运行，请从系统托盘图标打开。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new HomeworkViewer());

            mutex.ReleaseMutex();
        }
    }
}