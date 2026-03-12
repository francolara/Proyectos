
//////using DTO;
//////using Microsoft.ApplicationBlocks.Data;
using Dapper;
using ProyectoGRE.DTO;
using Microsoft.ApplicationBlocks.Data;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.IO;

namespace ProyectoGRE.DAO
{
    public class DaoGuiaRemision : DaoB
    {

        public Dto_GuiaRemision Lista_Consulta_Comp_GR(DtoB dtoBase)
        {
            Dto_GuiaRemision_Param dto = (Dto_GuiaRemision_Param)dtoBase;
            Dto_GuiaRemision cr = new Dto_GuiaRemision();
            SqlParameter[] pr = new SqlParameter[3];
            try
            {
                pr[0] = new SqlParameter("@Tipo", SqlDbType.Int);
                pr[0].Value = dto.Tipo;

                pr[1] = new SqlParameter("@Annio", SqlDbType.Int);
                pr[1].Value = 0;

                pr[2] = new SqlParameter("@Mes", SqlDbType.Int);
                pr[2].Value = 0;

                DataSet ds = SqlHelper.ExecuteDataset(objCn, CommandType.StoredProcedure, "Spu_Lista_GRElectronicos", pr);
                cr.DT = ds.Tables.Count >= 1 ? ds.Tables[0] : null;

            }
            catch (Exception ex)
            {
                cr.DT = null;
                cr.LugarError = ex.StackTrace;
                cr.ErrorEx = ex.Message;
                cr.ErrorMsj = "Error al Listar las Guias";
            }
            return cr;
        }

        public Dto_GuiaRemision Lista_GR(DtoB dtoBase)
        {
            Dto_GuiaRemision_Param dto = (Dto_GuiaRemision_Param)dtoBase;
            Dto_GuiaRemision cr = new Dto_GuiaRemision();
            SqlParameter[] pr = new SqlParameter[3];
            try
            {
                pr[0] = new SqlParameter("@Tipo", SqlDbType.Int);
                pr[0].Value = dto.Tipo;

                pr[1] = new SqlParameter("@Annio", SqlDbType.Int);
                pr[1].Value = 0;

                pr[2] = new SqlParameter("@Mes", SqlDbType.Int);
                pr[2].Value = 0;

                DataSet ds = SqlHelper.ExecuteDataset(objCn, CommandType.StoredProcedure, "Spu_Lista_GRElectronicos", pr);
                cr.DT = ds.Tables.Count >= 1 ? ds.Tables[0] : null;


            }
            catch (Exception ex)
            {
                cr.DT = null;
                cr.LugarError = ex.StackTrace;
                cr.ErrorEx = ex.Message;
                cr.ErrorMsj = "Error al Listar las Guias";
            }
            return cr;
        }

        public Dto_GuiaRemision Envia_GR_Det(DtoB dtoBase)
        {
            Dto_GuiaRemision_Param dto = (Dto_GuiaRemision_Param)dtoBase;
            Dto_GuiaRemision cr = new Dto_GuiaRemision();
            SqlParameter[] pr = new SqlParameter[5];
            try
            {
                pr[0] = new SqlParameter("@Tipo", SqlDbType.Int);
                pr[0].Value = 2;

                pr[1] = new SqlParameter("@TipDoc", SqlDbType.Char);
                pr[1].Value = "86";

                pr[2] = new SqlParameter("@SerDoc", SqlDbType.VarChar);
                pr[2].Value = dto.Serie;

                pr[3] = new SqlParameter("@NumDoc", SqlDbType.VarChar);
                pr[3].Value = dto.Numero;

                pr[4] = new SqlParameter("@Idempresa", SqlDbType.VarChar);
                pr[4].Value = dto.Idempresa;

                DataSet ds = SqlHelper.ExecuteDataset(objCn, CommandType.StoredProcedure, "Spu_Genera_Doc_electronico", pr);
                cr.DT_D = ds.Tables.Count >= 1 ? ds.Tables[0] : null;


                //if (pr[0].Value.ToString() != "")
                //{
                //    cr.LugarError = ToString("Lista_Gr_electro()");
                //    cr.ErrorMsj = pr[0].Value.ToString();
                //}
            }
            catch (Exception ex)
            {
                cr.DT_D = null;
                cr.LugarError = ex.StackTrace;
                cr.ErrorEx = ex.Message;
                cr.ErrorMsj = "Error al trae Detalle";
            }
            return cr;
        }
        public Dto_GuiaRemision Envia_GR(DtoB dtoBase)
        {
            Dto_GuiaRemision_Param dto = (Dto_GuiaRemision_Param)dtoBase;
            Dto_GuiaRemision cr = new Dto_GuiaRemision();
            SqlParameter[] pr = new SqlParameter[5];
            try
            {
                pr[0] = new SqlParameter("@Tipo", SqlDbType.Int);
                pr[0].Value = 1;

                pr[1] = new SqlParameter("@TipDoc", SqlDbType.Char);
                pr[1].Value = "86";

                pr[2] = new SqlParameter("@SerDoc", SqlDbType.VarChar);
                pr[2].Value = dto.Serie;

                pr[3] = new SqlParameter("@NumDoc", SqlDbType.VarChar);
                pr[3].Value = dto.Numero;

                pr[4] = new SqlParameter("@Idempresa", SqlDbType.VarChar);
                pr[4].Value = dto.Idempresa;


                DataSet ds = SqlHelper.ExecuteDataset(objCn, CommandType.StoredProcedure, "Spu_Genera_Doc_electronico", pr);
                cr.DT = ds.Tables.Count >= 1 ? ds.Tables[0] : null;

            }
            catch (Exception ex)
            {
                cr.DT = null;
                cr.LugarError = ex.StackTrace;
                cr.ErrorEx = ex.Message;
                cr.ErrorMsj = "Error al trae Cabecera";
            }
            return cr;
        }


        public Dto_GuiaRemision Act_Estado_GR(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            DtoRespuestaPSE_OSE_NUBEFACT dto = (DtoRespuestaPSE_OSE_NUBEFACT)_dtoRES;
            Dto_GuiaRemision cr = new Dto_GuiaRemision();
            SqlParameter[] pr = new SqlParameter[10];
            String CodError;
            try
            {
                if (dto.C_CAMPO_RESPUESTA_02 != null)
                {
                    CodError = "9"; /*TIENE UN ERROR VALIDADO POR NUBEFACT*/
                }
                else
                {
                    CodError = "3"; /*ACEPTADO POR NUBEFACT*/
                }

                pr[0] = new SqlParameter("@Tipo", SqlDbType.Int);
                pr[0].Value = 1;
                pr[1] = new SqlParameter("@TipDoc", SqlDbType.Char);
                pr[1].Value = dto.C_TIPO_DOCUMENTO_ORI;
                pr[2] = new SqlParameter("@SerDoc", SqlDbType.VarChar);
                pr[2].Value = dto.C_SERIE_DOCUMENTO;
                pr[3] = new SqlParameter("@NumDoc", SqlDbType.VarChar);
                pr[3].Value = dto.C_NUMERO_DOCUMENTO_ORI;
                pr[4] = new SqlParameter("@IdEmpresa", SqlDbType.Char);
                pr[4].Value = dto.C_ID_EMPRESA;
                pr[5] = new SqlParameter("@CodError", SqlDbType.VarChar);
                pr[5].Value = CodError;
                pr[6] = new SqlParameter("@DesError", SqlDbType.VarChar);
                pr[6].Value = dto.C_CAMPO_RESPUESTA_02;
                pr[7] = new SqlParameter("@PDF", SqlDbType.VarChar);
                pr[7].Value = "";
                pr[8] = new SqlParameter("@CDR", SqlDbType.VarChar);
                pr[8].Value = "";
                pr[9] = new SqlParameter("@XML", SqlDbType.VarChar);
                pr[9].Value = "";


                DataSet ds = SqlHelper.ExecuteDataset(objCn, CommandType.StoredProcedure, "Spu_ActEstado_DocElectronicos", pr);
                cr.DT = ds.Tables.Count >= 1 ? ds.Tables[0] : null;
            }
            catch (Exception ex)
            {
                cr.DT = null;
                cr.LugarError = ex.StackTrace;
                cr.ErrorEx = ex.Message;
                cr.ErrorMsj = "Error al trae Cabecera";
            }
            return cr;
        }

        public Dto_GuiaRemision Consulta_Estado_GR(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            DtoRespuestaPSE_OSE_NUBEFACT dto = (DtoRespuestaPSE_OSE_NUBEFACT)_dtoRES;
            Dto_GuiaRemision cr = new Dto_GuiaRemision();
            SqlParameter[] pr = new SqlParameter[10];
            String CodError;
            String DesError;
            int Tipo;
            try
            {
                if (dto.C_CAMPO_REST_BIT_2) /* si es true*/
                {
                    Tipo = 7; /*ACEPTADO*/
                    CodError = "";
                    DesError = "";
                }
                else
                {
                    Tipo = 8; /*RECHAZADO*/
                    CodError = "9";
                    DesError = dto.C_CAMPO_RESPUESTA_15;
                }

                pr[0] = new SqlParameter("@Tipo", SqlDbType.Int);
                pr[0].Value = Tipo;
                pr[1] = new SqlParameter("@TipDoc", SqlDbType.Char);
                pr[1].Value = dto.C_TIPO_DOCUMENTO_ORI;
                pr[2] = new SqlParameter("@SerDoc", SqlDbType.VarChar);
                pr[2].Value = dto.C_SERIE_DOCUMENTO;
                pr[3] = new SqlParameter("@NumDoc", SqlDbType.VarChar);
                pr[3].Value = dto.C_NUMERO_DOCUMENTO_ORI;
                pr[4] = new SqlParameter("@IdEmpresa", SqlDbType.Char);
                pr[4].Value = dto.C_ID_EMPRESA;
                pr[5] = new SqlParameter("@CodError", SqlDbType.VarChar);
                pr[5].Value = CodError;
                pr[6] = new SqlParameter("@DesError", SqlDbType.VarChar);
                pr[6].Value = DesError;
                pr[7] = new SqlParameter("@PDF", SqlDbType.VarChar);
                pr[7].Value = dto.C_CAMPO_RESPUESTA_11;
                pr[8] = new SqlParameter("@CDR", SqlDbType.VarChar);
                pr[8].Value = dto.C_CAMPO_RESPUESTA_13;
                pr[9] = new SqlParameter("@XML", SqlDbType.VarChar);
                pr[9].Value = dto.C_CAMPO_RESPUESTA_12;

                if (DesError.ToUpper() == "EN PROCESO")
                {
                    return cr;
                }

                DataSet ds = SqlHelper.ExecuteDataset(objCn, CommandType.StoredProcedure, "Spu_ActEstado_DocElectronicos", pr);
                cr.DT = ds.Tables.Count >= 1 ? ds.Tables[0] : null;
            }
            catch (Exception ex)
            {
                cr.DT = null;
                cr.LugarError = ex.StackTrace;
                cr.ErrorEx = ex.Message;
                cr.ErrorMsj = "Error al trae Cabecera";
            }
            return cr;
        }
        public string ToString(string metodo)
        {
            return "\n\rClase error: " + base.ToString() + "\n\rMetodo error: " + metodo;
        }

    }
}
