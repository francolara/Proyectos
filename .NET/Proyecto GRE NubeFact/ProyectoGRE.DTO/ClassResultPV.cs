using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;

namespace ProyectoGRE.DTO
{
    [Serializable]
    public class ClassResultPV
    {
        string resultado;
        string errorEx;
        string errorMsj;
        string alertaMsj;
        string lugarError;
        DataTable dt;
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
        public List<T> ConvertToGenericList<T>()
        {
            if (list == null)
                list = new List<DtoB>();

            ArrayList arrayList = new ArrayList(list);
            return new List<T>(arrayList.ToArray(typeof(T)) as T[]);
        }


        [Browsable(false)]
        public string Resultado
        {
            get { return resultado; }
            set { resultado = value; }
        }
    }
}
