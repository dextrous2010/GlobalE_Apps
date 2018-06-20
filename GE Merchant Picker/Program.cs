using System;
using System.Threading;
using System.Windows.Forms;

namespace GE_Merchant_Picker
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            using (Mutex mutex = new Mutex(false, "Global\\" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name))
            {
                if (!mutex.WaitOne(0, false))
                {
                    MessageBox.Show("Instance of application already running!");
                    return;
                }
                GC.Collect();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new GE_Merchant_Picker_Form());
            }
        }

        //private static string appGuid = "c0a76b5a-12ab-45c5-b9d9-d693faa6e7b9";
    }
}
