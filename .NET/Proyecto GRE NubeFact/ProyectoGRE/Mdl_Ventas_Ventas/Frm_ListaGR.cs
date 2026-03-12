using ProyectoGRE.DAO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ProyectoGRE.DTO;
using ProyectoGRE.Controladora;

namespace ProyectoGRE
{
    public partial class Frm_ListaGR : Form
    {
        public Frm_ListaGR()
        {
            InitializeComponent();
        }

        private void Frm_ListaGR_Load(object sender, EventArgs e)
        {

            // Ocultar primero
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Hide();

            // NotifyIcon
            NTFNB.Visible = true;
            NTFNB.BalloonTipTitle = "Servicio activo";
            NTFNB.BalloonTipText = "La tarea está ejecutándose en segundo plano";
            NTFNB.ShowBalloonTip(3000);

            // Iniciar timer AL FINAL
            timer1.Start();

        }

        private void CmdEnvSunat_Click(object sender, EventArgs e)
        {

            Consultat_Enviar_DocFac();
            Consultat_Enviar_GR();
        }

        private void Frm_ListaGR_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                this.ShowInTaskbar = false;
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            Consultat_Enviar_DocFac();
            Consultat_Enviar_GR();
        }

        private void Consultat_Enviar_DocFac()
        {
                try
                {
                    CtrDocFacturacion ctr = new CtrDocFacturacion();
                    Dto_GuiaRemision_Param dto = new Dto_GuiaRemision_Param();

                    dto.Tipo = 4;
                    Dto_GuiaRemision crc = ctr.Lista_Consulta_Comp_DocFac(dto);
                    DtoRespuestaPSE_OSE_NUBEFACT _dtoRES = new DtoRespuestaPSE_OSE_NUBEFACT();

                    dto.Tipo = 3;
                    Dto_GuiaRemision cr = ctr.Lista_DocFac(dto);

                    /* ENVIAR COMPROBANTE */
                    Dgv_Guias.DataSource = cr.DT;
                    if (cr.HuboError)
                        MessageBox.Show(cr.ErrorMsj + cr.Detalle);
                    else
                    {

                        foreach (DataRow row in cr.DT.Rows)
                        {
                            dto.Idempresa = Convert.ToString(row["idEmpresa"]);
                            dto.Serie = Convert.ToString(row["SERIE"]);
                            dto.Numero = Convert.ToString(row["NumOri"]);
                            dto.TipoDoc = Convert.ToString(row["Tipo"]);
                        Dto_GuiaRemision cr2 = ctr.Envia_DocFac(dto);
                        }
                    }

                /* CONSULTAR ESTADO COMPROBANTE */
                if (crc.HuboError)
                    MessageBox.Show(crc.ErrorMsj + crc.Detalle);
                else
                {

                    foreach (DataRow row in crc.DT.Rows)
                    {

                        _dtoRES.C_TIPO_DOCUMENTO = Convert.ToString(row["TIPD"]);
                        _dtoRES.C_TIPO_DOCUMENTO_ORI = Convert.ToString(row["Tipo"]);
                        _dtoRES.C_SERIE_DOCUMENTO = Convert.ToString(row["SERIE"]);
                        _dtoRES.C_NUMERO_DOCUMENTO = Convert.ToString(row["NUMDOC"]);
                        _dtoRES.C_NUMERO_DOCUMENTO_ORI = Convert.ToString(row["numeroori"]);
                        _dtoRES.C_ID_EMPRESA = Convert.ToString(row["idEmpresa"]);
                        Dto_GuiaRemision crc2 = ctr.Consultar_Comp_DocFac(_dtoRES);
                    }
                }

                }
                catch (Exception ex)
                {

                    //MessageBox (ex.StackTrace, ex.Message, "Error al cargar Guias");
                    MessageBox.Show(ex.Message);
                }

        }
        private void Consultat_Enviar_GR()
        {

                try
                {
                    CtrGuiaRemision ctr = new CtrGuiaRemision();
                    Dto_GuiaRemision_Param dto = new Dto_GuiaRemision_Param();

                    dto.Tipo = 2;
                    Dto_GuiaRemision crc = ctr.Lista_Consulta_Comp_GR(dto);
                    DtoRespuestaPSE_OSE_NUBEFACT _dtoRES = new DtoRespuestaPSE_OSE_NUBEFACT();

                    dto.Tipo = 1;
                    Dto_GuiaRemision cr = ctr.Lista_GR(dto);

                    /* ENVIAR COMPROBANTE */
                    Dgv_Guias.DataSource = cr.DT;
                    if (cr.HuboError)
                        MessageBox.Show(cr.ErrorMsj + cr.Detalle);
                    else
                    {

                        foreach (DataRow row in cr.DT.Rows)
                        {
                            dto.Idempresa = Convert.ToString(row["idEmpresa"]);
                            dto.Serie = Convert.ToString(row["SERIE"]);
                            dto.Numero = Convert.ToString(row["NumOri"]);
                            dto.TipoDoc = Convert.ToString(row["Tipo"]);
                            Dto_GuiaRemision cr2 = ctr.Envia_GR(dto);
                        }
                    }

                /* CONSULTAR ESTADO COMPROBANTE */
                if (crc.HuboError)
                    MessageBox.Show(crc.ErrorMsj + crc.Detalle);
                else
                {

                    foreach (DataRow row in crc.DT.Rows)
                    {

                        _dtoRES.C_TIPO_DOCUMENTO = Convert.ToString(row["TIPD"]);
                        _dtoRES.C_TIPO_DOCUMENTO_ORI = Convert.ToString(row["Tipo"]);
                        _dtoRES.C_SERIE_DOCUMENTO = Convert.ToString(row["SERIE"]);
                        _dtoRES.C_NUMERO_DOCUMENTO = Convert.ToString(row["NUMDOC"]);
                        _dtoRES.C_NUMERO_DOCUMENTO_ORI = Convert.ToString(row["numeroori"]);
                        _dtoRES.C_ID_EMPRESA = Convert.ToString(row["idEmpresa"]);
                        Dto_GuiaRemision crc2 = ctr.Consultar_Comp_GR(_dtoRES);
                    }
                }

                }
                catch (Exception ex)
                {

                    //MessageBox (ex.StackTrace, ex.Message, "Error al cargar Guias");
                    MessageBox.Show(ex.Message);
                }

        }

        private void salirToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            NTFNB.Visible = false;
            timer1.Stop();
            Application.Exit();
        }

        private void ejecutarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Consultat_Enviar_DocFac();
            Consultat_Enviar_GR();
        }
    }
}
