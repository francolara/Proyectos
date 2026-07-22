using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SistemaVisual.Models;
using SistemaVisual.Services;

namespace SistemaVisual
{
    // =============================================
    // Author:        FRANCO LARA
    // Create date:   21/07/2026
    // Description:   Interfaz asíncrona para instalación completa, reanudación y actualización desde Cloudflare R2.
    // =============================================
    public sealed class MainForm : Form
    {
        private readonly Label etiquetaTitulo;
        private readonly Label etiquetaEstado;
        private readonly Label etiquetaArchivo;
        private readonly Label etiquetaPorcentaje;
        private readonly Label etiquetaConteo;
        private readonly Label etiquetaResumen;
        private readonly ProgressBar barraProgreso;
        private readonly ProgressBar indicadorActividad;
        private readonly Button botonReintentar;
        private readonly Button botonLogs;
        private CancellationTokenSource cancelacion;
        private UpdateService servicioActual;
        private bool permitirCierre;
        private bool instalandoArchivos;
        private bool ejecutando;

        public MainForm()
        {
            Text = "SistemaVisual";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = true;
            ClientSize = new Size(650, 430);
            BackColor = Color.FromArgb(246, 248, 252);
            Font = new Font("Segoe UI", 9F);

            var encabezado = new Panel { Dock = DockStyle.Top, Height = 92, BackColor = Color.FromArgb(22, 78, 99) };
            etiquetaTitulo = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI Semibold", 20F, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(34, 18),
                Text = "Actualizando sistema..."
            };
            var subtitulo = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5F),
                ForeColor = Color.FromArgb(205, 232, 240),
                Location = new Point(37, 59),
                Text = "Los archivos se validan antes de incorporarse al sistema"
            };
            encabezado.Controls.Add(etiquetaTitulo);
            encabezado.Controls.Add(subtitulo);

            indicadorActividad = new ProgressBar
            {
                Location = new Point(38, 114),
                Size = new Size(574, 7),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 28
            };
            etiquetaEstado = new Label
            {
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Location = new Point(38, 143),
                Size = new Size(574, 28),
                Text = "Preparando actualizador..."
            };
            etiquetaArchivo = new Label
            {
                AutoEllipsis = true,
                ForeColor = Color.FromArgb(75, 85, 99),
                Location = new Point(38, 178),
                Size = new Size(574, 24),
                Text = " "
            };
            barraProgreso = new ProgressBar
            {
                Location = new Point(38, 211),
                Size = new Size(504, 23),
                Minimum = 0,
                Maximum = 100
            };
            etiquetaPorcentaje = new Label
            {
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(22, 78, 99),
                Location = new Point(548, 211),
                Size = new Size(64, 23),
                TextAlign = ContentAlignment.MiddleRight,
                Text = "0%"
            };
            etiquetaConteo = new Label
            {
                ForeColor = Color.FromArgb(75, 85, 99),
                Location = new Point(38, 251),
                Size = new Size(574, 22),
                Text = "Archivos procesados: 0 de 0"
            };
            etiquetaResumen = new Label
            {
                ForeColor = Color.FromArgb(55, 65, 81),
                Location = new Point(38, 279),
                Size = new Size(574, 25),
                Text = "Nuevos: 0   Actualizados: 0   Sin cambios: 0   Omitidos: 0"
            };

            botonReintentar = CrearBoton("Volver a intentar", new Point(338, 335), 132);
            botonLogs = CrearBoton("Abrir logs", new Point(480, 335), 132);
            botonReintentar.Visible = false;
            botonLogs.Visible = false;
            botonReintentar.Click += async (s, e) => await EjecutarActualizadorAsync();
            botonLogs.Click += (s, e) => AbrirLogs();

            var nota = new Label
            {
                ForeColor = Color.FromArgb(107, 114, 128),
                Location = new Point(38, 392),
                Size = new Size(574, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Text = "No apague el equipo mientras se estén reemplazando archivos."
            };

            Controls.Add(encabezado);
            Controls.Add(indicadorActividad);
            Controls.Add(etiquetaEstado);
            Controls.Add(etiquetaArchivo);
            Controls.Add(barraProgreso);
            Controls.Add(etiquetaPorcentaje);
            Controls.Add(etiquetaConteo);
            Controls.Add(etiquetaResumen);
            Controls.Add(botonReintentar);
            Controls.Add(botonLogs);
            Controls.Add(nota);

            Shown += async (s, e) => await EjecutarActualizadorAsync();
            FormClosing += MainForm_FormClosing;
        }

        private static Button CrearBoton(string texto, Point ubicacion, int ancho)
        {
            return new Button
            {
                Text = texto,
                Location = ubicacion,
                Size = new Size(ancho, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(22, 78, 99),
                Cursor = Cursors.Hand
            };
        }

        private async Task EjecutarActualizadorAsync()
        {
            if (ejecutando) return;
            ejecutando = true;
            permitirCierre = false;
            botonReintentar.Visible = false;
            botonLogs.Visible = false;
            indicadorActividad.MarqueeAnimationSpeed = 28;
            cancelacion = new CancellationTokenSource();
            LogService log = null;

            try
            {
                var configService = new ConfigService();
                bool configuracionCreada;
                var configuracion = configService.CargarOCrear(out configuracionCreada);
                if (configuracionCreada)
                    throw new InvalidOperationException("Se creó actualizador.config.json junto al ejecutable. Configure UrlWorker y vuelva a intentar.");

                var rutas = new AppPaths(configuracion);
                log = new LogService(rutas.ArchivoLog);
                servicioActual = new UpdateService(configuracion, rutas, log);
                var resultado = await servicioActual.EjecutarAsync(
                    new Progress<UpdateProgress>(ActualizarInterfaz),
                    SolicitarCierreVentasAsync,
                    cancelacion.Token);

                etiquetaEstado.Text = resultado.Mensaje;
                etiquetaResumen.Text = string.Format("Nuevos: {0}   Actualizados: {1}   Sin cambios: {2}   Omitidos: {3}",
                    resultado.Nuevos, resultado.Actualizados, resultado.SinCambio, resultado.Omitidos);
                await Task.Delay(1400, cancelacion.Token);
                etiquetaEstado.Text = "Iniciando sistema...";
                servicioActual.IniciarSistema();
                CerrarAplicacion();
            }
            catch (OperationCanceledException)
            {
                if (!IsDisposed)
                    MostrarError("La operación fue cancelada.");
            }
            catch (UnauthorizedAccessException ex)
            {
                if (log != null) log.Error("Falta de permisos de escritura.", ex);
                MostrarError("No se pudo escribir en C:\\Sistema Visual. Verifique los permisos de administrador.");
            }
            catch (IOException ex)
            {
                if (log != null) log.Error("Error de archivos durante la operación.", ex);
                MostrarError("No se pudo completar la operación. Verifique el espacio disponible y que los archivos no estén en uso.");
            }
            catch (Exception ex)
            {
                if (log != null) log.Error("Error presentado a la interfaz.", ex);
                MostrarError(ex.Message);
            }
            finally
            {
                ejecutando = false;
                if (cancelacion != null) cancelacion.Dispose();
                cancelacion = null;
            }
        }

        private Task<bool> SolicitarCierreVentasAsync()
        {
            var respuesta = MessageBox.Show(this,
                "Ventas.exe está abierto. Guarde su trabajo y presione Reintentar para solicitar un cierre normal.\n\n"
                + "El actualizador no finalizará el proceso por la fuerza.",
                "Sistema en uso", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
            return Task.FromResult(respuesta == DialogResult.Retry);
        }

        private void ActualizarInterfaz(UpdateProgress avance)
        {
            instalandoArchivos = avance.Instalando;
            etiquetaTitulo.Text = avance.ModoInstalacion ? "Instalando Sistema Visual..." : "Actualizando sistema...";
            etiquetaEstado.Text = string.IsNullOrWhiteSpace(avance.Estado) ? "Procesando..." : avance.Estado;
            etiquetaArchivo.Text = string.IsNullOrWhiteSpace(avance.Archivo) ? " " : avance.Archivo;
            etiquetaPorcentaje.Text = avance.Porcentaje + "%";
            etiquetaConteo.Text = string.Format("Archivos procesados: {0} de {1}", avance.Procesados, avance.Total);
            etiquetaResumen.Text = string.Format("Nuevos: {0}   Actualizados: {1}   Sin cambios: {2}   Omitidos: {3}",
                avance.Nuevos, avance.Actualizados, avance.SinCambios, avance.Omitidos);
            barraProgreso.Style = avance.Indeterminado ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            barraProgreso.MarqueeAnimationSpeed = avance.Indeterminado ? 24 : 0;
            if (!avance.Indeterminado)
                barraProgreso.Value = Math.Max(0, Math.Min(100, avance.Porcentaje));
        }

        private void MostrarError(string mensaje)
        {
            if (IsDisposed) return;
            instalandoArchivos = false;
            permitirCierre = true;
            indicadorActividad.MarqueeAnimationSpeed = 0;
            etiquetaTitulo.Text = "No se pudo completar el proceso";
            etiquetaEstado.Text = mensaje;
            botonReintentar.Visible = true;
            botonLogs.Visible = servicioActual != null;
            MessageBox.Show(this, mensaje, "SistemaVisual", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void AbrirLogs()
        {
            try
            {
                if (servicioActual != null) servicioActual.AbrirCarpetaLogs();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo abrir la carpeta de logs: " + ex.Message,
                    "SistemaVisual", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CerrarAplicacion()
        {
            instalandoArchivos = false;
            permitirCierre = true;
            Close();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (permitirCierre) return;
            if (instalandoArchivos)
            {
                e.Cancel = true;
                System.Media.SystemSounds.Beep.Play();
                return;
            }
            if (!ejecutando || e.CloseReason != CloseReason.UserClosing) return;

            var respuesta = MessageBox.Show(this, "¿Desea cancelar el proceso y cerrar?",
                "Confirmar cierre", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (respuesta != DialogResult.Yes)
                e.Cancel = true;
            else if (cancelacion != null)
                cancelacion.Cancel();
        }
    }
}
