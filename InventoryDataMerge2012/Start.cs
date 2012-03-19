using System;
using System.Windows.Forms;

namespace InventoryDataMerge2012
{
    static class Start
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new UserWBookWSheet());
        }
    }
}
