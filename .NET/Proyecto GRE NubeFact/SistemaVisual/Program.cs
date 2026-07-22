using System;
using System.Threading;
using System.Windows.Forms;

namespace SistemaVisual
{
    internal static class Program
    {
        private const string NombreMutex = @"Global\FRALSETECH_SistemaVisual_Actualizador";

        [STAThread]
        private static void Main()
        {
            bool mutexCreado;
            using (var mutex = new Mutex(true, NombreMutex, out mutexCreado))
            {
                if (!mutexCreado)
                {
                    MessageBox.Show(
                        "El actualizador ya se encuentra ejecutándose.",
                        "Sistema Visual",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
                GC.KeepAlive(mutex);
            }
        }
    }
}
