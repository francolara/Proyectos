using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace ProyectoGRE.DTO
{


    public class DtoRespuestaPSE_OSE_NUBEFACT
    {
        public String C_ERROR { get; set; }
        public String C_JSON_COMPROBANTE { get; set; }
        public String C_TOKEN_ENVIO { get; set; }
        public String C_JSON_ENVIO { get; set; }
        public String C_NOMBRE_ARCHIVO { get; set; }
        public String C_RUC_EMPRESA { get; set; }
        public String C_RUC_CLIENTE{ get; set; }
        public DateTime C_FECHA_ENVIO { get; set; }
        public String C_TIPO_DOCUMENTO_ORI { get; set; }
        public String C_TIPO_DOCUMENTO { get; set; }
        public String C_ID_EMPRESA { get; set; }
        public String C_TIPO_DOCUMENTO_ANUL { get; set; }
        public String C_SERIE_DOCUMENTO { get; set; }
        public String C_NUMERO_DOCUMENTO { get; set; }
        public String C_NUMERO_DOCUMENTO_ORI { get; set; }
        public String C_COD_VENTA { get; set; }
        public Boolean C_IB_ANULADO { get; set; }
        public Boolean C_IB_RESUMEN { get; set; }
        public Boolean C_IB_ES_CONSULTA { get; set; }
        public String C_CAMPO_RESPUESTA_01 { get; set; }
        public String C_CAMPO_RESPUESTA_02 { get; set; }
        public String C_CAMPO_RESPUESTA_03 { get; set; }
        public String C_CAMPO_RESPUESTA_04 { get; set; }
        public String C_CAMPO_RESPUESTA_05 { get; set; }
        public String C_CAMPO_RESPUESTA_06 { get; set; }
        public String C_CAMPO_RESPUESTA_07 { get; set; }
        public String C_CAMPO_RESPUESTA_08 { get; set; }
        public String C_CAMPO_RESPUESTA_09 { get; set; }
        public String C_CAMPO_RESPUESTA_10 { get; set; }
        public String C_CAMPO_RESPUESTA_11 { get; set; }
        public String C_CAMPO_RESPUESTA_12 { get; set; }
        public String C_CAMPO_RESPUESTA_13 { get; set; }
        public String C_CAMPO_RESPUESTA_14 { get; set; }
        public String C_CAMPO_RESPUESTA_15 { get; set; }
        public String C_CAMPO_RESPUESTA_16 { get; set; }
        public String C_CAMPO_RESPUESTA_17 { get; set; }
        public String C_CAMPO_RESPUESTA_18 { get; set; }
        public String C_CAMPO_RESPUESTA_19 { get; set; }
        public String C_CAMPO_RESPUESTA_20 { get; set; }
        public String C_CAMPO_RUTA_DOC { get; set; }
        public String C_CAMPO_RUTA_DOC_XML { get; set; }
        public String C_CAMPO_RUTA_1 { get; set; }
        public String C_CAMPO_RUTA_2 { get; set; }
        public String C_CAMPO_RUTA_3 { get; set; }
        public String C_CAMPO_RUTA_4 { get; set; }

        public Boolean C_CAMPO_REST_BIT_1 { get; set; }
        public Boolean C_CAMPO_REST_BIT_2 { get; set; }
        public Boolean C_CAMPO_REST_BIT_3 { get; set; }
        public Boolean C_CAMPO_REST_BIT_4 { get; set; }
        public Boolean C_CAMPO_REST_BIT_5 { get; set; }
    }

    public class Dto_GuiaRemision_Param : DtoB
    {

        string cd_Prf;
        string prueba1;
        int prueba2;
        string descrip;
        string nomP;
        string nvl_prf;
        string serie;
        string idempresa;
        string tipoDoc;
        string numero;
        bool estado;
        int tipo;
        int tipoCab;
        int tipoDet;

        public int Tipo
        {
            get { return tipo; }
            set { tipo = value; }
        }

        public int TipoCab
        {
            get { return tipoCab; }
            set { tipoCab = value; }
        }

        public int TipoDet
        {
            get { return tipoDet; }
            set { tipoDet = value; }
        }
        public string TipoDoc
        {
            get { return tipoDoc; }
            set { tipoDoc = value; }
        }


        public string Serie
        {
            get { return serie; }
            set { serie = value; }
        }

        public string Idempresa
        {
            get { return idempresa; }
            set { idempresa = value; }
        }
        
        public string Numero
        {
            get { return numero; }
            set { numero = value; }
        }

        public string Prueba1
        {
            get { return prueba1; }
            set { prueba1 = value; }
        }
        /// <summary> Codigo Perfil --> nvarchar(3)</summary>
        public int Prueba2
        {
            get { return prueba2; }
            set { prueba2 = value; }
        }

        /// <summary> Descripcion --> varchar(200),	null</summary>
        public string Descrip
        {
            get { return descrip; }
            set { descrip = value; }
        }

        /// <summary> Nombre Perfil -->	varchar(30)</summary>
        public string NomP
        {
            get { return nomP; }
            set { nomP = value; }
        }

        /// <summary> Estado --> bit </summary>
        public bool Estado
        {
            get { return estado; }
            set { estado = value; }
        }

        /// <summary> Nvl_Prf ---> varchar(100) </summary>
        public string Nvl_Prf
        {
            get { return nvl_prf; }
            set { nvl_prf = value; }
        }

        private string cd_MN;

        public string Cd_MN
        {
            get { return cd_MN; }
            set { cd_MN = value; }
        }

        private bool iB_TPermisos;

        public bool IB_TPermisos
        {
            get { return iB_TPermisos; }
            set { iB_TPermisos = value; }
        }

        //Auxiliares
        private string _nomUsu;
        public string NomUsu { get { return _nomUsu; } set { _nomUsu = value; } }

    }

        public class Dto_GuiaRemision : DtoB
    {
        string resultado;
        string errorEx;
        string errorMsj;
        string alertaMsj;
        string lugarError;
        DataTable dt;
        DataTable dt_d;
        DataSet ds;
        List<DtoB> list;
        DtoB entidad;
        object objeto;

        [Browsable(false)]
        public object Objeto
        {
            get { return objeto; }
            set { objeto = value; }
        }

        /// <summary>
        /// lista de  dto's!!  creada por pepito! para k no moleste juan pin! (GET or SET)
        /// </summary>
        public List<DtoB> List
        {
            get { return list; }
            set { list = value; }
        }
        [Browsable(false)]
        public DtoB Entidad
        {
            get { return entidad; }
            set { entidad = value; }
        }
        bool huboError;

        bool huboAlerta;

        /// <summary>Tabla de resultado (GET or SET) </summary>
        [Browsable(false)]
        public DataTable DT
        {
            get { return dt; }
            set { dt = value; }
        }

        public DataTable DT_D
        {
            get { return dt_d; }
            set { dt_d = value; }
        }

        public DataTable DT_GR
        {
            get { return dt_d; }
            set { dt_d = value; }
        }

        /// <summary>DataSet de resultado (GET or SET) </summary>
        [Browsable(false)]
        public DataSet DS
        {
            get { return ds; }
            set { ds = value; }
        }

        /// <summary>Nos indica si hubo algun error (GET) </summary>
        [Browsable(false)]
        public bool HuboError
        {
            get { return huboError; }
            set { huboError = value; }
        }

        /// <summary>Nos indica si hubo alguna alerta (GET) </summary>
        [Browsable(false)]
        public bool HuboAlerta
        {
            get { return huboAlerta; }
            set { huboAlerta = value; }
        }

        /// <summary>Lugar de donde se produjo el error (SET) </summary>
        [Browsable(false)]
        public string LugarError
        {
            get { return lugarError; }
            set { lugarError = value; }
        }

        /// <summary>Msj de error que vera el Usuario (GET or SET) </summary>
        [Browsable(false)]
        public string ErrorMsj
        {
            get { return errorMsj; }
            set
            {
                errorMsj = value;
                huboError = true;
            }
        }

        /// <summary>Msj de alerta que vera el Usuario (GET or SET) </summary>
        [Browsable(false)]
        public string AlertaMsj
        {
            get { return alertaMsj; }
            set
            {
                alertaMsj = value;
                huboAlerta = true;
            }
        }

        /// <summary>Msj de error que provoca la Excepcion (SET) </summary>
        [Browsable(false)]
        public string ErrorEx
        {
            get { return errorEx; }
            set { errorEx = value; }
        }

        /// <summary>Nos muestra el detalle del error (GET) </summary>
        [Browsable(false)]
        public string Detalle
        {
            get { return errorEx + "\n\rLUGAR DE ERROR :" + lugarError; }
        }

        public string NombreArchivo { get; set; }

        public string NombreArchivoCdr { get; set; }

        /// <summary>
        /// Metodo de conversion entre listas (DtoB a T)
        /// </summary>
        /// <typeparam name="T">Tipo de Dto a la cual se convertira la lista</typeparam>
        /// <returns></returns>
        //////public List<T> ConvertToGenericList<T>()
        //////{
        //////    if (list == null)
        //////        list = new List<DtoB>();

        //////    ArrayList arrayList = new ArrayList(list);
        //////    return new List<T>(arrayList.ToArray(typeof(T)) as T[]);
        //////}


        [Browsable(false)]
        public string Resultado
        {
            get { return resultado; }
            set { resultado = value; }
        }

    }
}
