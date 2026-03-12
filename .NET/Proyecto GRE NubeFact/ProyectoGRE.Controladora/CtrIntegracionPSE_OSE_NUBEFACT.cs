//using Controladora.A_Utilitarios;
//using DAO.IntegracionPSE_OSE_SUNAT;
using ProyectoGRE.DTO;
//using DTO.IntegracionPSE_OSE_NUBEFACT;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using RestSharp;
using System.Data;
using ProyectoGRE.Controladora;

namespace Controladora.IntegracionPSE_OSE_SUNAT
{
    public class CtrIntegracionPSE_OSE_NUBEFACT 
    {

        #region Constructor



        #endregion

        #region Metodos

        private void Carga_Datos_Principales()
        {
            //try
            //{
            //    _dtoCONF = (DtoConfiguracionPseOse)_ctrCONF.USP_T_CONFIGURACION_PSE_OSE_CONSULTAR(_dtoCONF).Entidad;
            //    _listaRUTAS = _daoNUBF.Consulta_Rutas_Servidor_PSE_OSE(Utilitarios.RucEmp);
            //}
            //catch (Exception)
            //{ }
        }

        public string Enviar_PSE_OSE_NUBEFACT(Dto_GuiaRemision DtoGrCab, Dto_GuiaRemision DtoGrdet)
        {
            string error = String.Empty;
            string JSON = String.Empty;
            DtoRespuestaPSE_OSE_NUBEFACT _dtoRES = new DtoRespuestaPSE_OSE_NUBEFACT();

            try
            {
                Crear_JSON_GUIA_REMISION(_dtoRES, DtoGrCab.DT, DtoGrdet.DT_D);
                Crear_Json_ENVIO_COMPROBANTE(_dtoRES);
           
                /*ENVIAMOS DOCUMENTO A LA OSE*/
                Enviar_Json_OSE(_dtoRES);
            }
            catch (Exception bz)
            {
                error = bz.Message;
            }
            return error;
        }


        public string Enviar_PSE_OSE_NUBEFACT_DocFac(Dto_GuiaRemision DtoGrCab, Dto_GuiaRemision DtoGrdet, Dto_GuiaRemision DtoGrdet_GR)
        {
            string error = String.Empty;
            string JSON = String.Empty;
            DtoRespuestaPSE_OSE_NUBEFACT _dtoRES = new DtoRespuestaPSE_OSE_NUBEFACT();

            try
            {
                Crear_JSON_DocFac(_dtoRES, DtoGrCab.DT, DtoGrdet.DT_D, DtoGrdet_GR.DT_GR);
                Crear_Json_ENVIO_COMPROBANTE(_dtoRES);

                /*ENVIAMOS DOCUMENTO A LA OSE*/
                Enviar_Json_OSE_DocFac(_dtoRES);
            }
            catch (Exception bz)
            {
                error = bz.Message;
            }
            return error;
        }

        public string CONSULTA_PSE_OSE_NUBEFACT(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            string error = String.Empty;
            string JSON = String.Empty;
            DtoRespuestaPSE_OSE_NUBEFACT _dtoRES = new DtoRespuestaPSE_OSE_NUBEFACT();
            try
            {
                Crear_Json_CONSULTA_COMPROBANTE(_dtoENV);
                Consultar_Json_OSE(_dtoENV);
            }
            catch (Exception bz)
            {
                error = bz.Message;
            }
            return error;
        }

        public string CONSULTA_PSE_OSE_NUBEFACT_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            string error = String.Empty;
            string JSON = String.Empty;
            DtoRespuestaPSE_OSE_NUBEFACT _dtoRES = new DtoRespuestaPSE_OSE_NUBEFACT();

            try
            {
                Crear_Json_CONSULTA_COMPROBANTE_DocFac(_dtoENV);
                Consultar_Json_OSE_DocFac(_dtoENV);
            }
            catch (Exception bz)
            {
                error = bz.Message;
            }
            return error;
        }

        public void Crear_Json_CONSULTA_COMPROBANTE(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            try
            {                
                var _DIC_S_O = new Dictionary<string, object>();

                _DIC_S_O.Add("operacion", "consultar_guia");
                _DIC_S_O.Add("tipo_de_comprobante", Convert.ToString(!String.IsNullOrWhiteSpace(_dtoENV.C_TIPO_DOCUMENTO) ? (_dtoENV.C_TIPO_DOCUMENTO == "09" ? "7" : "0") : ""));
                _DIC_S_O.Add("serie", _dtoENV.C_SERIE_DOCUMENTO);
                _DIC_S_O.Add("numero", Convert.ToString(Convert.ToInt32(_dtoENV.C_NUMERO_DOCUMENTO)));

                _dtoENV.C_JSON_ENVIO = JsonConvert.SerializeObject(_DIC_S_O);
            }
            catch (Exception bz)
            {
                _dtoENV.C_ERROR = bz.Message;
            }
        }

        public void Crear_Json_CONSULTA_COMPROBANTE_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            try
            {

                var _DIC_S_O = new Dictionary<string, object>();

                _DIC_S_O.Add("operacion", "consultar_comprobante");
                _DIC_S_O.Add("tipo_de_comprobante", _dtoENV.C_TIPO_DOCUMENTO);
                _DIC_S_O.Add("serie", _dtoENV.C_SERIE_DOCUMENTO);
                _DIC_S_O.Add("numero", Convert.ToString(Convert.ToInt32(_dtoENV.C_NUMERO_DOCUMENTO)));

                _dtoENV.C_JSON_ENVIO = JsonConvert.SerializeObject(_DIC_S_O);
            }
            catch (Exception bz)
            {
                _dtoENV.C_ERROR = bz.Message;
            }
        }

        public void Crear_Json_ENVIO_COMPROBANTE(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            try
            {
                var utf8WithoutBom = new UTF8Encoding(false);

                string json = JsonConvert.SerializeObject(_dtoENV.C_JSON_COMPROBANTE, Formatting.Indented);

                byte[] bytes = Encoding.Default.GetBytes(json);
                _dtoENV.C_JSON_ENVIO = _dtoENV.C_JSON_COMPROBANTE;
            }
            catch (Exception bz)
            {
                _dtoENV.C_ERROR = bz.Message;
            }
        }

        public void Crear_JSON_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV, DataTable _Comprobante, DataTable _Comprobant_Det, DataTable _Comprobant_Det_GR)

        {
            try
            {
                Dictionary<string, object> _DIC_S_O = new Dictionary<string, object>();
                Dictionary<string, string> _DIC_S_S = new Dictionary<string, string>();
                List<Dictionary<string, string>> _LST_DIC_S_s = new List<Dictionary<string, string>>();

                #region Comprobante

                _DIC_S_O.Add("operacion", "generar_comprobante");
                _DIC_S_O.Add("tipo_de_comprobante", _Comprobante.Rows[0]["Tipo_comprobante"].ToString());
                _DIC_S_O.Add("serie", _Comprobante.Rows[0]["SERIE"].ToString());
                _DIC_S_O.Add("numero", Convert.ToString(_Comprobante.Rows[0]["NUMDOC"].ToString()));
                _DIC_S_O.Add("sunat_transaction", _Comprobante.Rows[0]["TipoOpe_nb"].ToString());
                _DIC_S_O.Add("cliente_tipo_de_documento", _Comprobante.Rows[0]["tipo_cli"].ToString());
                _DIC_S_O.Add("cliente_numero_de_documento", _Comprobante.Rows[0]["Ruc"].ToString());
                _DIC_S_O.Add("cliente_denominacion", _Comprobante.Rows[0]["Rzn"].ToString());
                _DIC_S_O.Add("cliente_direccion", _Comprobante.Rows[0]["direccion"].ToString());
                _DIC_S_O.Add("cliente_email", "");
                _DIC_S_O.Add("cliente_email_1", "");
                _DIC_S_O.Add("cliente_email_2", "");

                _DIC_S_O.Add("fecha_de_emision", _Comprobante.Rows[0]["FecDocEmi"].ToString());
                _DIC_S_O.Add("fecha_de_vencimiento", _Comprobante.Rows[0]["Fecvcto"].ToString());
                _DIC_S_O.Add("moneda", _Comprobante.Rows[0]["Moneda"].ToString());
                _DIC_S_O.Add("tipo_de_cambio", _Comprobante.Rows[0]["TipoCambio"].ToString());
                _DIC_S_O.Add("porcentaje_de_igv", _Comprobante.Rows[0]["Igv"].ToString());

                _DIC_S_O.Add("descuento_global", "");
                _DIC_S_O.Add("total_descuento", "");
                _DIC_S_O.Add("total_anticipo", "");
                _DIC_S_O.Add("total_gravada", _Comprobante.Rows[0]["TotalValorVenta"].ToString());
                _DIC_S_O.Add("total_inafecta", "");
                _DIC_S_O.Add("total_exonerada", "");
                _DIC_S_O.Add("total_igv", _Comprobante.Rows[0]["TotalIGVVenta"].ToString());

                _DIC_S_O.Add("total_gratuita", "");
                _DIC_S_O.Add("total_otros_cargos", "");
                _DIC_S_O.Add("total", _Comprobante.Rows[0]["TotalPrecioVenta"].ToString());

                //if (_Comprobante.Rows[0]["PorcPercep"].ToString() != "0.00%")
                //{
                //    _DIC_S_O.Add("percepcion_tipo", "1");
                //    _DIC_S_O.Add("percepcion_base_imponible", _Comprobante.Rows[0]["TotBasePercep"].ToString());
                //    _DIC_S_O.Add("total_percepcion", _Comprobante.Rows[0]["TotPercep"].ToString());

                //    decimal totBasePercep = _Comprobante.Rows[0]["TotBasePercep"] == DBNull.Value ? 0 : Convert.ToDecimal(_Comprobante.Rows[0]["TotBasePercep"]);
                //    decimal totPercep = _Comprobante.Rows[0]["TotPercep"] == DBNull.Value ? 0 : Convert.ToDecimal(_Comprobante.Rows[0]["TotPercep"]);
                //    decimal totalIncluidoPercepcion = totBasePercep + totPercep;

                //    _DIC_S_O.Add("total_incluido_percepcion", totalIncluidoPercepcion.ToString("0.00"));
                //}
                //else
                //{
                _DIC_S_O.Add("percepcion_tipo", "");
                _DIC_S_O.Add("percepcion_base_imponible", "");
                _DIC_S_O.Add("total_percepcion", "");
                _DIC_S_O.Add("total_incluido_percepcion", "");
                //}

                if (_Comprobante.Rows[0]["AgenteRet"].ToString() == "1")
                {
                    _DIC_S_O.Add("retencion_tipo", _Comprobante.Rows[0]["TipoRet"].ToString());
                    _DIC_S_O.Add("retencion_base_imponible", _Comprobante.Rows[0]["TotalPrecioVenta"].ToString());
                    _DIC_S_O.Add("total_retencion", _Comprobante.Rows[0]["TotRet"].ToString());
                }
                else
                {
                    _DIC_S_O.Add("retencion_tipo", "");
                    _DIC_S_O.Add("retencion_base_imponible", "");
                    _DIC_S_O.Add("total_retencion", "");
                }

                _DIC_S_O.Add("total_impuestos_bolsas", "");
                _DIC_S_O.Add("observaciones", _Comprobante.Rows[0]["ObsDocVentas"].ToString());

                if (_Comprobante.Rows[0]["CodDetrac"].ToString() == "")
                {
                    _DIC_S_O.Add("detraccion", "false");

                }
                else
                {
                    _DIC_S_O.Add("detraccion", "true");

                }

                if (_Comprobante.Rows[0]["idDocumento"].ToString() == "07")
                {
                    _DIC_S_O.Add("documento_que_se_modifica_tipo", _Comprobante.Rows[0]["TipDoc_Mod"].ToString());
                    _DIC_S_O.Add("documento_que_se_modifica_serie", _Comprobante.Rows[0]["SerDoc_Mod"].ToString());
                    _DIC_S_O.Add("documento_que_se_modifica_numero", _Comprobante.Rows[0]["NumDoc_Mod"].ToString());
                    _DIC_S_O.Add("tipo_de_nota_de_credito", _Comprobante.Rows[0]["Tiponc"].ToString());
                    _DIC_S_O.Add("tipo_de_nota_de_debito", "");
                }
                else
                {
                    _DIC_S_O.Add("documento_que_se_modifica_tipo", "");
                    _DIC_S_O.Add("documento_que_se_modifica_serie", "");
                    _DIC_S_O.Add("documento_que_se_modifica_numero", "");
                    _DIC_S_O.Add("tipo_de_nota_de_credito", "");
                    _DIC_S_O.Add("tipo_de_nota_de_debito", "");
                }
                _DIC_S_O.Add("enviar_automaticamente_a_la_sunat", "true");
                _DIC_S_O.Add("enviar_automaticamente_al_cliente", "false");
                _DIC_S_O.Add("codigo_unico", "");

                _DIC_S_O.Add("condiciones_de_pago", _Comprobante.Rows[0]["glsFormaPago"].ToString());
                _DIC_S_O.Add("medio_de_pago", _Comprobante.Rows[0]["GlsMedioPago"].ToString());
                _DIC_S_O.Add("placa_vehiculo", "");
                _DIC_S_O.Add("orden_compra_servicio", "");

                if (_Comprobante.Rows[0]["CodDetrac"].ToString() != "")
                {
                    _DIC_S_O.Add("detraccion_tipo", _Comprobante.Rows[0]["CodDetrac"].ToString());
                    _DIC_S_O.Add("detraccion_total", _Comprobante.Rows[0]["TotDetrac"].ToString());
                    _DIC_S_O.Add("medio_de_pago_detraccion", "1");
                }

                _DIC_S_O.Add("formato_de_pdf", "A4");
                _DIC_S_O.Add("generado_por_contingencia", "");
                _DIC_S_O.Add("bienes_region_selva", "");
                _DIC_S_O.Add("servicios_region_selva", "");

                #endregion

                #region Detalle

                var LST_DETALLE = new List<Dictionary<string, object>>();

                foreach (DataRow row in _Comprobant_Det.Rows)
                {
                    var DETALLE = new Dictionary<string, object>();

                    DETALLE.Add("unidad_de_medida", Convert.ToString(row["UM"]));
                    DETALLE.Add("codigo", Convert.ToString(row["idProducto"]));
                    DETALLE.Add("descripcion", Convert.ToString(row["GlsProducto"]));
                    DETALLE.Add("cantidad", Convert.ToString(row["Cantidad"]));
                    DETALLE.Add("valor_unitario", Convert.ToString(row["VVUnit"]));
                    DETALLE.Add("precio_unitario", Convert.ToString(row["PVUnit"]));
                    DETALLE.Add("descuento", "");
                    DETALLE.Add("subtotal", Convert.ToString(row["TotalVVNeto"]));
                    DETALLE.Add("tipo_de_igv", Convert.ToString(row["tipo_igv"]));
                    DETALLE.Add("igv", Convert.ToString(row["TotalIGVNeto"]));
                    DETALLE.Add("total", Convert.ToString(row["TotalPVNeto"]));
                    DETALLE.Add("anticipo_regularizacion", "false");
                    DETALLE.Add("anticipo_documento_serie", "");
                    DETALLE.Add("anticipo_documento_numero", "");

                    LST_DETALLE.Add(DETALLE);
                }

                _DIC_S_O.Add("items", LST_DETALLE);
                #endregion

                #region Guias
                var LST_DETALLE_GUIAS = new List<Dictionary<string, object>>();

                foreach (DataRow row in _Comprobant_Det_GR.Rows)
                {
                    var DETALLE_G = new Dictionary<string, object>();

                    DETALLE_G.Add("guia_tipo", "1");
                    DETALLE_G.Add("guia_serie_numero", Convert.ToString(row["Numguia"]));

                    LST_DETALLE_GUIAS.Add(DETALLE_G);
                }

                _DIC_S_O.Add("guias", LST_DETALLE_GUIAS);
                #endregion

                #region Cuotas

                if (_Comprobante.Rows[0]["GlsMedioPago"].ToString() == "Crédito")
                {

                    if (_Comprobante.Rows[0]["idDocumento"].ToString() == "01" || _Comprobante.Rows[0]["idDocumento"].ToString() == "07")
                    {

                        var LST_DETALLE_CUOTAS = new List<Dictionary<string, object>>();
                        var DETALLE_C = new Dictionary<string, object>();

                        DETALLE_C.Add("cuota", "1");
                        if (_Comprobante.Rows[0]["idDocumento"].ToString() == "07" && _Comprobante.Rows[0]["Tiponc"].ToString() == "13")
                        {
                            DETALLE_C.Add("fecha_de_pago", _Comprobante.Rows[0]["Fecvcto"].ToString());
                        }
                        else
                        {
                            DETALLE_C.Add("fecha_de_pago", _Comprobante.Rows[0]["Fecvcto"].ToString());
                        }

                        decimal totTotDetrac_nb = _Comprobante.Rows[0]["TotDetrac"] == DBNull.Value ? 0 : Convert.ToDecimal(_Comprobante.Rows[0]["TotDetrac"]);
                        decimal totMtoTot = _Comprobante.Rows[0]["TotalPrecioVenta"] == DBNull.Value ? 0 : Convert.ToDecimal(_Comprobante.Rows[0]["TotalPrecioVenta"]);
                        decimal totRetencion = _Comprobante.Rows[0]["TotRet"] == DBNull.Value ? 0 : Convert.ToDecimal(_Comprobante.Rows[0]["TotRet"]);
                        decimal total = totMtoTot - (totTotDetrac_nb + totRetencion);

                        DETALLE_C.Add("importe", total.ToString("0.00"));

                        LST_DETALLE_CUOTAS.Add(DETALLE_C);

                        _DIC_S_O.Add("venta_al_credito", LST_DETALLE_CUOTAS);
                    }
                }
                #endregion
                
                /*ARMAR JSON*/
                _dtoENV.C_JSON_COMPROBANTE = JsonConvert.SerializeObject(_DIC_S_O);
                //_dtoENV.C_NOMBRE_ARCHIVO = _Comprobante.C_NOMBRE_ARCHIVO;
                _dtoENV.C_TIPO_DOCUMENTO = Convert.ToString(_Comprobante.Rows[0]["idDocumento"]);
                _dtoENV.C_TIPO_DOCUMENTO_ORI = Convert.ToString(_Comprobante.Rows[0]["idDocumento"]);
                _dtoENV.C_SERIE_DOCUMENTO = Convert.ToString(_Comprobante.Rows[0]["SERIE"]);
                _dtoENV.C_NUMERO_DOCUMENTO = Convert.ToString(_Comprobante.Rows[0]["NUMDOC"]);
                _dtoENV.C_NUMERO_DOCUMENTO_ORI = Convert.ToString(_Comprobante.Rows[0]["NumDocOri"]);
                _dtoENV.C_ID_EMPRESA = Convert.ToString(_Comprobante.Rows[0]["idEmpresa"]);
            }
            catch (Exception bz)
            {
                //_dtoENV.C_ERROR = bz.Message;
            }
        }


        public void Crear_JSON_GUIA_REMISION(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV,DataTable _Comprobante, DataTable _Comprobant_Det)

        {
            try
            {
                Dictionary<string, object> _DIC_S_O = new Dictionary<string, object>();
                Dictionary<string, string> _DIC_S_S = new Dictionary<string, string>();
                List<Dictionary<string, string>> _LST_DIC_S_s = new List<Dictionary<string, string>>();

                #region Comprobante

                _DIC_S_O.Add("operacion", "generar_guia");
                _DIC_S_O.Add("tipo_de_comprobante", 7);
                _DIC_S_O.Add("serie", _Comprobante.Rows[0]["SERIE"].ToString());
                _DIC_S_O.Add("numero", Convert.ToString(_Comprobante.Rows[0]["NUMDOC"].ToString()));
                _DIC_S_O.Add("cliente_tipo_de_documento", _Comprobante.Rows[0]["tipo_cli"].ToString());
                _DIC_S_O.Add("cliente_numero_de_documento", _Comprobante.Rows[0]["Ruc"].ToString());
                _DIC_S_O.Add("cliente_denominacion", _Comprobante.Rows[0]["Rzn"].ToString());
                _DIC_S_O.Add("cliente_direccion", _Comprobante.Rows[0]["CLI_DIROFI"].ToString());
                _DIC_S_O.Add("cliente_email", "");
                _DIC_S_O.Add("cliente_email_1", "");
                _DIC_S_O.Add("cliente_email_2", "");
                
                _DIC_S_O.Add("fecha_de_emision", _Comprobante.Rows[0]["FecDocEmi"].ToString());
                _DIC_S_O.Add("observaciones", _Comprobante.Rows[0]["ObsDocVentas"].ToString());

                _DIC_S_O.Add("motivo_de_traslado", _Comprobante.Rows[0]["Motivo_Tras"].ToString()); 
                _DIC_S_O.Add("peso_bruto_total", _Comprobante.Rows[0]["PESOBRUTO"].ToString());
                _DIC_S_O.Add("peso_bruto_unidad_de_medida", "KGM");

                _DIC_S_O.Add("numero_de_bultos", _Comprobante.Rows[0]["Bultos"].ToString());
                _DIC_S_O.Add("tipo_de_transporte", _Comprobante.Rows[0]["Tipo_Tras"].ToString());
                _DIC_S_O.Add("fecha_de_inicio_de_traslado", _Comprobante.Rows[0]["FecIniTraslado"].ToString());
               
                _DIC_S_O.Add("transportista_documento_tipo", _Comprobante.Rows[0]["TIPODOCTRAN"].ToString());
                _DIC_S_O.Add("transportista_documento_numero", _Comprobante.Rows[0]["DNIRUCTRA"].ToString());
                _DIC_S_O.Add("transportista_denominacion", _Comprobante.Rows[0]["TRANOM"].ToString());
                _DIC_S_O.Add("transportista_placa_numero", _Comprobante.Rows[0]["PLACA"].ToString());

                _DIC_S_O.Add("tuc_vehiculo_principal", _Comprobante.Rows[0]["tuc_vehiculo_principal"].ToString());
                _DIC_S_O.Add("conductor_documento_tipo", _Comprobante.Rows[0]["conductor_documento_tipo"].ToString());
                _DIC_S_O.Add("conductor_documento_numero", _Comprobante.Rows[0]["conductor_documento_numero"].ToString());
                _DIC_S_O.Add("conductor_denominacion", _Comprobante.Rows[0]["conductor_denominacion"].ToString());
                _DIC_S_O.Add("conductor_nombre", _Comprobante.Rows[0]["conductor_nombre"].ToString());
                _DIC_S_O.Add("conductor_apellidos", _Comprobante.Rows[0]["conductor_apellidos"].ToString());
                _DIC_S_O.Add("conductor_numero_licencia", _Comprobante.Rows[0]["conductor_numero_licencia"].ToString());
                _DIC_S_O.Add("destinatario_documento_tipo", _Comprobante.Rows[0]["destinatario_documento_tipo"].ToString());
                _DIC_S_O.Add("destinatario_documento_numero", _Comprobante.Rows[0]["destinatario_documento_numero"].ToString());
                _DIC_S_O.Add("destinatario_denominacion", _Comprobante.Rows[0]["destinatario_denominacion"].ToString());
                _DIC_S_O.Add("mtc", _Comprobante.Rows[0]["mtc"].ToString());


                _DIC_S_O.Add("punto_de_partida_ubigeo", _Comprobante.Rows[0]["AL1_Ubigeo_Partida"].ToString());
                _DIC_S_O.Add("punto_de_partida_direccion", _Comprobante.Rows[0]["Partida"].ToString());
                _DIC_S_O.Add("punto_de_partida_codigo_establecimiento_sunat", _Comprobante.Rows[0]["punto_de_partida_codigo_establecimiento_sunat"].ToString());
                _DIC_S_O.Add("punto_de_llegada_ubigeo", _Comprobante.Rows[0]["AL1_Ubigeo_Llegada"].ToString());
                _DIC_S_O.Add("punto_de_llegada_direccion", _Comprobante.Rows[0]["llegada"].ToString());
                _DIC_S_O.Add("punto_de_llegada_codigo_establecimiento_sunat", _Comprobante.Rows[0]["punto_de_llegada_codigo_establecimiento_sunat"].ToString());


                _DIC_S_O.Add("enviar_automaticamente_al_cliente", "true");
                _DIC_S_O.Add("formato_de_pdf", "");

                #endregion

                #region Detalle

                var LST_DETALLE = new List<Dictionary<string, object>>();

                foreach (DataRow row in _Comprobant_Det.Rows)
                {
                    var DETALLE = new Dictionary<string, object>();

                    DETALLE.Add("unidad_de_medida", Convert.ToString(row["UM"]));
                    DETALLE.Add("codigo", Convert.ToString(row["idProducto"]));
                    DETALLE.Add("descripcion", Convert.ToString(row["GlsProducto"]));
                    DETALLE.Add("cantidad", Convert.ToString(row["Cantidad"]));

                    LST_DETALLE.Add(DETALLE);
                }

                _DIC_S_O.Add("items", LST_DETALLE);
                #endregion

                /*ARMAR JSON*/
                _dtoENV.C_JSON_COMPROBANTE = JsonConvert.SerializeObject(_DIC_S_O);
                //_dtoENV.C_NOMBRE_ARCHIVO = _Comprobante.C_NOMBRE_ARCHIVO;

                _dtoENV.C_TIPO_DOCUMENTO = Convert.ToString(_Comprobante.Rows[0]["TIPDOC"]);
                _dtoENV.C_TIPO_DOCUMENTO_ORI = Convert.ToString(_Comprobante.Rows[0]["idDocumento"]);
                _dtoENV.C_SERIE_DOCUMENTO = Convert.ToString(_Comprobante.Rows[0]["SERIE"]);
                _dtoENV.C_NUMERO_DOCUMENTO = Convert.ToString(_Comprobante.Rows[0]["NUMDOC"]);
                _dtoENV.C_NUMERO_DOCUMENTO_ORI = Convert.ToString(_Comprobante.Rows[0]["NumDocOri"]);
                _dtoENV.C_ID_EMPRESA = Convert.ToString(_Comprobante.Rows[0]["idEmpresa"]);

            }
            catch (Exception bz)
            {
                //_dtoENV.C_ERROR = bz.Message;
            }
        }
        public void Enviar_Json_OSE(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            try
            {

                CtrGuiaRemision ctr = new CtrGuiaRemision();

                string url = String.Empty;
                String Token = String.Empty;

                //pruebas
                url = "https://api.nubefact.com/api/v1/5a2eb047-d278-4708-a8b0-5b6d0c8d2637";
                Token = "4933554f554a41c7b4007b40d8b50165640a97989edc44118f94d5c6903cee75";

                ////Produccion
                //url = "";
                //Token = "";

                string _respuesta = String.Empty;

                if (!String.IsNullOrEmpty(url))
                    _respuesta = sendPOST(Token, url, _dtoRES.C_JSON_ENVIO);    

                dynamic _jsonRespuesta = JsonConvert.DeserializeObject(_respuesta);
                //_dtoRES.C_RUC_EMPRESA = "20503141389";
                //_dtoRES.C_FECHA_ENVIO = DateTime.Now;

                foreach (var item in _jsonRespuesta)
                {
                    if (!_dtoRES.C_IB_ANULADO)
                    {
                        if (_respuesta.Contains("errors"))
                        {
                            if (item.Name.Equals("codigo"))
                                _dtoRES.C_CAMPO_RESPUESTA_01 = item.Value;

                            if (item.Name.Equals("errors"))
                                _dtoRES.C_CAMPO_RESPUESTA_02 = item.Value;
                        }
                        else
                        {

                            if (item.Name.Equals("enlace_del_pdf"))
                                _dtoRES.C_CAMPO_RESPUESTA_11 = item.Value;

                            if (item.Name.Equals("enlace_del_xml"))
                                _dtoRES.C_CAMPO_RESPUESTA_12 = item.Value;

                            if (item.Name.Equals("enlace_del_cdr"))
                                _dtoRES.C_CAMPO_RESPUESTA_13 = item.Value;
                        }
                    }
                    else
                    {
                        if (_respuesta.Contains("errors"))
                        {
                            if (item.Name.Equals("codigo"))
                                _dtoRES.C_CAMPO_RESPUESTA_14 = item.Value;

                            if (item.Name.Equals("errors"))
                                _dtoRES.C_CAMPO_RESPUESTA_15 = item.Value;
                        }
                        else
                        {

                            if (item.Name.Equals("enlace_del_pdf"))
                                _dtoRES.C_CAMPO_RESPUESTA_11 = item.Value;

                            if (item.Name.Equals("enlace_del_xml"))
                                _dtoRES.C_CAMPO_RESPUESTA_12 = item.Value;

                            if (item.Name.Equals("enlace_del_cdr"))
                                _dtoRES.C_CAMPO_RESPUESTA_13 = item.Value;
                        }
                    }
                }

                Dto_GuiaRemision cr = ctr.Act_Estado_GR(_dtoRES);
            }
            catch (Exception bz)
            {
                _dtoRES.C_ERROR = bz.Message;
            }
        }

        public void Enviar_Json_OSE_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            try
            {

                CtrDocFacturacion ctr = new CtrDocFacturacion();

                string url = String.Empty;
                String Token = String.Empty;

                //pruebas
                url = "https://api.nubefact.com/api/v1/5a2eb047-d278-4708-a8b0-5b6d0c8d2637";
                Token = "4933554f554a41c7b4007b40d8b50165640a97989edc44118f94d5c6903cee75";

                //Produccion
                //url = "";
                //Token = "4933554f554a41c7b4007b40d8b50165640a97989edc44118f94d5c6903cee75";
                string _respuesta = String.Empty;

                if (!String.IsNullOrEmpty(url))
                    _respuesta = sendPOST(Token, url, _dtoRES.C_JSON_ENVIO);

                dynamic _jsonRespuesta = JsonConvert.DeserializeObject(_respuesta);
                //_dtoRES.C_RUC_EMPRESA = "20503141389";
                //_dtoRES.C_FECHA_ENVIO = DateTime.Now;

                foreach (var item in _jsonRespuesta)
                {
                    if (!_dtoRES.C_IB_ANULADO)
                    {
                        if (_respuesta.Contains("errors"))
                        {
                            if (item.Name.Equals("codigo"))
                                _dtoRES.C_CAMPO_RESPUESTA_01 = item.Value;

                            if (item.Name.Equals("errors"))
                                _dtoRES.C_CAMPO_RESPUESTA_02 = item.Value;
                        }
                        else
                        {

                            if (item.Name.Equals("enlace_del_pdf"))
                                _dtoRES.C_CAMPO_RESPUESTA_11 = item.Value;

                            if (item.Name.Equals("enlace_del_xml"))
                                _dtoRES.C_CAMPO_RESPUESTA_12 = item.Value;

                            if (item.Name.Equals("enlace_del_cdr"))
                                _dtoRES.C_CAMPO_RESPUESTA_13 = item.Value;
                        }
                    }
                    else
                    {
                        if (_respuesta.Contains("errors"))
                        {
                            if (item.Name.Equals("codigo"))
                                _dtoRES.C_CAMPO_RESPUESTA_14 = item.Value;

                            if (item.Name.Equals("errors"))
                                _dtoRES.C_CAMPO_RESPUESTA_15 = item.Value;
                        }
                        else
                        {

                            if (item.Name.Equals("enlace_del_pdf"))
                                _dtoRES.C_CAMPO_RESPUESTA_11 = item.Value;

                            if (item.Name.Equals("enlace_del_xml"))
                                _dtoRES.C_CAMPO_RESPUESTA_12 = item.Value;

                            if (item.Name.Equals("enlace_del_cdr"))
                                _dtoRES.C_CAMPO_RESPUESTA_13 = item.Value;
                        }
                    }
                }

                Dto_GuiaRemision cr = ctr.Act_Estado_DocFac(_dtoRES);
            }
            catch (Exception bz)
            {
                _dtoRES.C_ERROR = bz.Message;
            }
        }

        public void Consultar_Json_OSE_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoCONS)
        {
            try
            {
                CtrDocFacturacion ctr = new CtrDocFacturacion();

                string url = String.Empty;
                String Token = String.Empty;

                //pruebas
                url = "https://api.nubefact.com/api/v1/5a2eb047-d278-4708-a8b0-5b6d0c8d2637";
                Token = "4933554f554a41c7b4007b40d8b50165640a97989edc44118f94d5c6903cee75";

                string _respuesta = String.Empty;

                if (!String.IsNullOrEmpty(url))
                    _respuesta = sendPOST(Token, url, _dtoCONS.C_JSON_ENVIO);

                dynamic _jsonRespuesta = JsonConvert.DeserializeObject(_respuesta);

                _dtoCONS.C_RUC_EMPRESA = "20503141389";
                _dtoCONS.C_FECHA_ENVIO = DateTime.Now;

                foreach (var item in _jsonRespuesta)
                {
                    if (_dtoCONS.C_IB_ES_CONSULTA)
                    {
                        if (item.Name.Equals("sunat_responsecode"))
                            _dtoCONS.C_CAMPO_RESPUESTA_01 = item.Value;

                        if (item.Name.Equals("sunat_description"))
                            _dtoCONS.C_CAMPO_RESPUESTA_02 = item.Value;

                        if (item.Name.Equals("sunat_ticket_numero"))
                            _dtoCONS.C_CAMPO_RESPUESTA_03 = item.Value;

                        if (item.Name.Equals("sunat_note"))
                            _dtoCONS.C_CAMPO_RESPUESTA_04 = item.Value;

                        if (item.Name.Equals("sunat_soap_error"))
                            _dtoCONS.C_CAMPO_RESPUESTA_05 = item.Value;

                        if (item.Name.Equals("enlace"))
                            _dtoCONS.C_CAMPO_RESPUESTA_06 = item.Value;

                        if (item.Name.Equals("cadena_para_codigo_qr"))
                            _dtoCONS.C_CAMPO_RESPUESTA_07 = item.Value;

                        if (item.Name.Equals("codigo_hash"))
                            _dtoCONS.C_CAMPO_RESPUESTA_08 = item.Value;

                        if (item.Name.Equals("codigo_de_barras"))
                            _dtoCONS.C_CAMPO_RESPUESTA_09 = item.Value;

                        if (item.Name.Equals("key"))
                            _dtoCONS.C_CAMPO_RESPUESTA_10 = item.Value;

                        if (item.Name.Equals("enlace_del_pdf"))
                            _dtoCONS.C_CAMPO_RESPUESTA_11 = item.Value;

                        if (item.Name.Equals("enlace_del_xml"))
                            _dtoCONS.C_CAMPO_RESPUESTA_12 = item.Value;

                        if (item.Name.Equals("enlace_del_cdr"))
                            _dtoCONS.C_CAMPO_RESPUESTA_13 = item.Value;

                        if (item.Name.Equals("aceptada_por_sunat"))
                            _dtoCONS.C_CAMPO_REST_BIT_1 = item.Value;
                    }
                    else
                    {
                        if (item.Name.Equals("sunat_responsecode"))
                            _dtoCONS.C_CAMPO_RESPUESTA_14 = item.Value;

                        if (item.Name.Equals("sunat_description"))
                            _dtoCONS.C_CAMPO_RESPUESTA_15 = item.Value;

                        if (item.Name.Equals("sunat_ticket_numero"))
                            _dtoCONS.C_CAMPO_RESPUESTA_03 = item.Value;

                        if (item.Name.Equals("sunat_note"))
                            _dtoCONS.C_CAMPO_RESPUESTA_16 = item.Value;

                        if (item.Name.Equals("sunat_soap_error"))
                            _dtoCONS.C_CAMPO_RESPUESTA_17 = item.Value;

                        if (item.Name.Equals("enlace"))
                            _dtoCONS.C_CAMPO_RESPUESTA_06 = item.Value;

                        if (item.Name.Equals("key"))
                            _dtoCONS.C_CAMPO_RESPUESTA_10 = item.Value;

                        if (item.Name.Equals("enlace_del_pdf"))
                            _dtoCONS.C_CAMPO_RESPUESTA_11 = item.Value;

                        if (item.Name.Equals("enlace_del_xml"))
                            _dtoCONS.C_CAMPO_RESPUESTA_12 = item.Value;

                        if (item.Name.Equals("enlace_del_cdr"))
                            _dtoCONS.C_CAMPO_RESPUESTA_13 = item.Value;

                        if (item.Name.Equals("aceptada_por_sunat"))
                            _dtoCONS.C_CAMPO_REST_BIT_2 = item.Value;
                    }
                }


                Dto_GuiaRemision cr = ctr.Consulta_Estado_DocFac(_dtoCONS);

            }
            catch (Exception bz)
            {
                _dtoCONS.C_ERROR = bz.Message;
            }
        }

        public void Consultar_Json_OSE(DtoRespuestaPSE_OSE_NUBEFACT _dtoCONS)
        {
            try
            {
                CtrGuiaRemision ctr = new CtrGuiaRemision();

                string url = String.Empty;
                String Token = String.Empty;
                //url = "https://api.nubefact.com/api/v1/c449cafe-ed26-46de-bd7f-4be354bd6c21";
                //Token = "efe6f0c6d506477c802bbbf10327ddc8d265ddface7746f78be209c3b0f379aa";

                //Pruebas
                //url = "https://api.nubefact.com/api/v1/8a667928-f1bc-4edb-9959-13d3018c3c2f";
                //Token = "57f8bc4db1444ece994555a8efa0903e55e1980000d54775811b3cddf2649ea4";
 
                //Produccion
                url = "https://api.nubefact.com/api/v1/5a2eb047-d278-4708-a8b0-5b6d0c8d2637";
                Token = "4933554f554a41c7b4007b40d8b50165640a97989edc44118f94d5c6903cee75";


                string _respuesta = String.Empty;

                if (!String.IsNullOrEmpty(url))
                    _respuesta = sendPOST(Token, url, _dtoCONS.C_JSON_ENVIO);

                dynamic _jsonRespuesta = JsonConvert.DeserializeObject(_respuesta);

                _dtoCONS.C_RUC_EMPRESA = "20503141389";
                _dtoCONS.C_FECHA_ENVIO = DateTime.Now;

                foreach (var item in _jsonRespuesta)
                {
                    if (_dtoCONS.C_IB_ES_CONSULTA)
                    {
                        if (item.Name.Equals("sunat_responsecode"))
                            _dtoCONS.C_CAMPO_RESPUESTA_01 = item.Value;

                        if (item.Name.Equals("sunat_description"))
                            _dtoCONS.C_CAMPO_RESPUESTA_02 = item.Value;

                        if (item.Name.Equals("sunat_ticket_numero"))
                            _dtoCONS.C_CAMPO_RESPUESTA_03 = item.Value;

                        if (item.Name.Equals("sunat_note"))
                            _dtoCONS.C_CAMPO_RESPUESTA_04 = item.Value;

                        if (item.Name.Equals("sunat_soap_error"))
                            _dtoCONS.C_CAMPO_RESPUESTA_05 = item.Value;

                        if (item.Name.Equals("enlace"))
                            _dtoCONS.C_CAMPO_RESPUESTA_06 = item.Value;

                        if (item.Name.Equals("cadena_para_codigo_qr"))
                            _dtoCONS.C_CAMPO_RESPUESTA_07 = item.Value;

                        if (item.Name.Equals("codigo_hash"))
                            _dtoCONS.C_CAMPO_RESPUESTA_08 = item.Value;

                        if (item.Name.Equals("codigo_de_barras"))
                            _dtoCONS.C_CAMPO_RESPUESTA_09 = item.Value;

                        if (item.Name.Equals("key"))
                            _dtoCONS.C_CAMPO_RESPUESTA_10 = item.Value;

                        if (item.Name.Equals("enlace_del_pdf"))
                            _dtoCONS.C_CAMPO_RESPUESTA_11 = item.Value;

                        if (item.Name.Equals("enlace_del_xml"))
                            _dtoCONS.C_CAMPO_RESPUESTA_12 = item.Value;

                        if (item.Name.Equals("enlace_del_cdr"))
                            _dtoCONS.C_CAMPO_RESPUESTA_13 = item.Value;

                        if (item.Name.Equals("aceptada_por_sunat"))
                            _dtoCONS.C_CAMPO_REST_BIT_1 = item.Value;
                    }
                    else
                    {
                        if (item.Name.Equals("sunat_responsecode"))
                            _dtoCONS.C_CAMPO_RESPUESTA_14 = item.Value;

                        if (item.Name.Equals("sunat_description"))
                            _dtoCONS.C_CAMPO_RESPUESTA_15 = item.Value;

                        if (item.Name.Equals("sunat_ticket_numero"))
                            _dtoCONS.C_CAMPO_RESPUESTA_03 = item.Value;

                        if (item.Name.Equals("sunat_note"))
                            _dtoCONS.C_CAMPO_RESPUESTA_16 = item.Value;

                        if (item.Name.Equals("sunat_soap_error"))
                            _dtoCONS.C_CAMPO_RESPUESTA_17 = item.Value;

                        if (item.Name.Equals("enlace"))
                            _dtoCONS.C_CAMPO_RESPUESTA_06 = item.Value;

                        if (item.Name.Equals("key"))
                            _dtoCONS.C_CAMPO_RESPUESTA_10 = item.Value;

                        if (item.Name.Equals("enlace_del_pdf"))
                            _dtoCONS.C_CAMPO_RESPUESTA_11 = item.Value;

                        if (item.Name.Equals("enlace_del_xml"))
                            _dtoCONS.C_CAMPO_RESPUESTA_12 = item.Value;

                        if (item.Name.Equals("enlace_del_cdr"))
                            _dtoCONS.C_CAMPO_RESPUESTA_13 = item.Value;

                        if (item.Name.Equals("aceptada_por_sunat"))
                            _dtoCONS.C_CAMPO_REST_BIT_2 = item.Value;
                    }
                }


                Dto_GuiaRemision cr = ctr.Consulta_Estado_GR(_dtoCONS);

            }
            catch (Exception bz)
            {
                _dtoCONS.C_ERROR = bz.Message;
            }
        }

        public String sendPOST(string token, string url, string json)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var client = new RestClient(url);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Authorization", "Bearer " + token);

                request.AddParameter("application/json", json, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                return response.Content;
            }
            catch (Exception bz)
            {
                return bz.Message;
            }
        }

        #endregion
    }
}
