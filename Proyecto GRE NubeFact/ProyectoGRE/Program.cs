using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace ProyectoGRE
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>

        private static Mutex mutex;
        [STAThread]
        static void Main()
        {
            bool createdNew;
            mutex = new Mutex(true, "SUNAT_API_GR_UNICA_INSTANCIA", out createdNew);

            if (!createdNew)
            {
                MessageBox.Show(
                    "La aplicación ya se encuentra en ejecución.\n" +
                    "Revise el ícono en la bandeja del sistema.",
                    "Servicio SUNAT",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                // Ya hay una instancia ejecutándose
                return;
            }

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Frm_ListaGR());

            mutex.ReleaseMutex();
        }
    }
}
