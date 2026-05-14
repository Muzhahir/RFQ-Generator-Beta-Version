using AutoUpdaterDotNET;
using RFQ_Generator_System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RFQ_Generator_Beta_Version
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            AutoUpdater.Mandatory = true;
            AutoUpdater.UpdateMode = Mode.ForcedDownload;
            AutoUpdater.CheckForUpdateEvent += (args) =>
            {
                if (args.Error != null)
                {
                    MessageBox.Show(args.Error.Message, "Update Error");
                }
            };
            AutoUpdater.Start("https://drive.google.com/uc?export=download&id=11cmu1EMPreEyTEzfQYhDb_uZG7hclqwj");

            Application.Run(new Form1()); 
        }
    }
}