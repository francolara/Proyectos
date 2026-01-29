using System;
using System.Collections.Generic;
using System.Text;

namespace ProyectoGRE.DTO
{
    public class DtoConexion
    {
        public string Nombre { get; set; }
        public string IpLocal { get; set; }
        public string IpPublico { get; set; }
        public string BD { get; set; }
        public string UserID { get; set; }
        public string Pass { get; set; }
        public bool Estado { get; set; }
        public bool UsaRemoto { get; set; }
        public bool EsConexionLocal { get; set; }
        public bool Lcn { get; set; }
        public string RucEmpresa { get; set; }

        public bool IsDirectorioRed { get; set; }
        public string ModoIngreso { get; set; }
        public string NombreIngreso { get; set; }
        public bool IsDefinicionLocal { get; set; }
        public string ProveedorOSE { get; set; }
        public bool CamposAdicionalesFE { get; set; }

    }
}
