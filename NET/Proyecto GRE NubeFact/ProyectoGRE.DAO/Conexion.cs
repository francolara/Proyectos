using ProyectoGRE.DAO;
using ProyectoGRE.DAO.Properties;
using ProyectoGRE.DTO;
using System;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Xml;
namespace DAO
{
    public class Conexion
    {
        private string strCon;

        public static bool EsConexionLocal { get; private set; }

        public static string IpServer { get; private set; }

        public static string Pass { get; private set; }

        public static string BD { get; set; }

        public static string UserID { get; set; }

        public static string Version { get; set; }

        public static bool Lcn { get; private set; }

        public static string RucEmpresa { get; set; }

        public static bool IsDirectorioRed { get; set; }
        public static string ModoIngreso { get; set; }
        public static string NombreIngreso { get; set; }
        public static bool IsDefinicionLocal { get; set; }
        public static string ProveedorOSE { get; set; }
        public static bool CamposAdicionalesFE { get; set; }
        public string StrCon
        {
            get { return strCon; }
            set { strCon = value; }
        }

        public static bool IncFileCnx { private get; set; }
        public static System.Collections.Generic.Dictionary<string, object> PrmFileCnx { private get; set; }

        public Conexion()
        {
            if (string.IsNullOrEmpty(DAO.Conexion.IpServer) || string.IsNullOrEmpty(DAO.Conexion.BD) || string.IsNullOrEmpty(DAO.Conexion.UserID))
            {
                if (!IncFileCnx)
                    CargarConexion(); // Version anterior aun activada
                else
                    CargarConexionFile();
            }

            strCon = "Data Source=" + IpServer + ";Initial Catalog=" + BD + "; User Id=" + UserID + "; Password=" + Pass + ";";
        }

        private void LeeConfiguracionConexion(DtoConexion dto)
        {
            StringReader _Reader = new StringReader(Resources.Configuraciones);
            XmlTextReader _xmlReader = new XmlTextReader(_Reader);

            XmlDocument xDoc = new XmlDocument();

            //La ruta del documento XML permite rutas relativas 
            //respecto del ejecutable!
            xDoc.Load(_xmlReader);

            XmlNodeList configuraciones = xDoc.GetElementsByTagName("Configuraciones");
            XmlNodeList lista = ((XmlElement)configuraciones[0]).GetElementsByTagName("Configuracion");

            foreach (XmlElement nodo in lista)
            {
                dto.Estado = Convert.ToBoolean(nodo.GetAttribute("Estado"));

                if (dto.Estado)
                {
                    dto.RucEmpresa = nodo.GetAttribute("RucEmpresa");
                    dto.Nombre = nodo.GetAttribute("Nombre");
                    dto.IpLocal = nodo.GetAttribute("IpLocal");
                    dto.IpPublico = nodo.GetAttribute("IpPublico");
                    dto.BD = nodo.GetAttribute("BD");
                    dto.UserID = nodo.GetAttribute("UserID");
                    dto.Pass = nodo.GetAttribute("Pass");

                    if (!nodo.GetAttribute("Lcn").Equals(""))
                        dto.Lcn = Convert.ToBoolean(nodo.GetAttribute("Lcn"));

                    if (!nodo.GetAttribute("Ext").Equals(""))
                        dto.IsDirectorioRed = Convert.ToBoolean(nodo.GetAttribute("Ext"));
                    if (!nodo.GetAttribute("ConexExt").Equals(""))
                        dto.ModoIngreso = nodo.GetAttribute("ConexExt");
                    if (!nodo.GetAttribute("RutaExt").Equals(""))
                        dto.NombreIngreso = nodo.GetAttribute("RutaExt");
                    if (!nodo.GetAttribute("ConexExtLocal").Equals(""))
                        dto.IsDefinicionLocal = Convert.ToBoolean(nodo.GetAttribute("ConexExtLocal"));

                    if (!nodo.GetAttribute("ProvOSE").Equals(""))
                        dto.ProveedorOSE = nodo.GetAttribute("ProvOSE");

                    if (!nodo.GetAttribute("CamAFE").Equals(""))
                        dto.CamposAdicionalesFE = Convert.ToBoolean(nodo.GetAttribute("CamAFE"));

                    break;
                }
            }
        }//fin Metodo

        private void CargarConexion()
        {
            DtoConexion dto = new DtoConexion();
            LeeConfiguracionConexion(dto);

            DAO.Conexion.IpServer = dto.IpLocal;
            EsConexionLocal = true;

            if (ValidaIP(dto.IpLocal))
            {
                DAO.Conexion.IpServer = dto.IpLocal;
                EsConexionLocal = true;
            }
            else
            {
                DAO.Conexion.IpServer = dto.IpPublico;
                EsConexionLocal = false;
            }
            DAO.Conexion.UserID = dto.UserID;
            DAO.Conexion.BD = dto.BD;
            DAO.Conexion.Pass = dto.Pass;
            DAO.Conexion.Lcn = dto.Lcn;
            DAO.Conexion.RucEmpresa = dto.RucEmpresa;

            IsDirectorioRed = dto.IsDirectorioRed;
            ModoIngreso = dto.ModoIngreso;
            NombreIngreso = dto.NombreIngreso;
            IsDefinicionLocal = dto.IsDefinicionLocal;
            ProveedorOSE = dto.ProveedorOSE;
            CamposAdicionalesFE = dto.CamposAdicionalesFE;
        }

        public static DtoConexion ObtenerConexion()
        {
            DtoConexion connect = new DtoConexion();
            connect.EsConexionLocal = Conexion.EsConexionLocal;
            connect.IpLocal = Conexion.IpServer;
            connect.IpPublico = Conexion.IpServer;
            connect.BD = Conexion.BD;
            connect.UserID = Conexion.UserID;
            connect.Pass = Conexion.Pass;
            connect.Lcn = Conexion.Lcn;
            connect.RucEmpresa = Conexion.RucEmpresa;

            connect.IsDirectorioRed = Conexion.IsDirectorioRed;
            connect.ModoIngreso = Conexion.ModoIngreso;
            connect.NombreIngreso = Conexion.NombreIngreso;
            connect.IsDefinicionLocal = Conexion.IsDefinicionLocal;
            connect.ProveedorOSE = Conexion.ProveedorOSE;
            connect.CamposAdicionalesFE = Conexion.CamposAdicionalesFE;
            return connect;
        }

        public static string ObtenerCadenaConexion()
        {
            return "Data Source=" + IpServer + ";Initial Catalog=" + BD + "; User Id=" + UserID + "; Password=" + Pass + ";";
        }

        private bool ValidaIP(string ip)
        {
            bool ping_out = false;
            try
            {
                IPAddress ip_address;
                Ping ping_ip = new Ping();
                PingReply pr;
                string status;
                if (ip.Contains(","))
                {
                    int i = ip.IndexOf(",");
                    ip = ip.Substring(0, i);
                }
                ip_address = IPAddress.Parse(ip);

                pr = ping_ip.Send(ip);
                status = pr.Status.ToString();
                if (status == IPStatus.Success.ToString())
                {
                    ping_out = true;
                }
            }
            catch (Exception)
            {
                ping_out = false;
            }
            return ping_out;
        }



        //Nuevo Proceso Para Leer Conexion
        private void CargarConexionFile()
        {
            DtoConexion dto = new DtoConexion();

            if (PrmFileCnx != null)
            {
                if (PrmFileCnx.ContainsKey("CodCli")) RucEmpresa = PrmFileCnx["CodCli"].ToString();

                DAO.Conexion.IpServer = PrmFileCnx["IpLocal"].ToString();
                EsConexionLocal = true;

                if (ValidaIP(PrmFileCnx["IpLocal"].ToString()))
                {
                    DAO.Conexion.IpServer = PrmFileCnx["IpLocal"].ToString();
                    EsConexionLocal = true;
                }
                else
                {
                    DAO.Conexion.IpServer = PrmFileCnx["IpPublico"].ToString();
                    EsConexionLocal = false;
                }

                DAO.Conexion.UserID = PrmFileCnx["UserID"].ToString();
                DAO.Conexion.BD = PrmFileCnx["BD"].ToString();
                DAO.Conexion.Pass = PrmFileCnx["Pass"].ToString();
                DAO.Conexion.Lcn = Convert.ToBoolean(PrmFileCnx["Lcn"].ToString());

                if (PrmFileCnx.ContainsKey("Ext")) IsDirectorioRed = Convert.ToBoolean(PrmFileCnx["Ext"].ToString());
                if (PrmFileCnx.ContainsKey("CnxExt")) ModoIngreso = PrmFileCnx["CnxExt"].ToString();
                if (PrmFileCnx.ContainsKey("RtaExt")) NombreIngreso = PrmFileCnx["RtaExt"].ToString();
                if (PrmFileCnx.ContainsKey("LocExt")) IsDefinicionLocal = Convert.ToBoolean(PrmFileCnx["LocExt"].ToString());
                if (PrmFileCnx.ContainsKey("ProvOSE")) ProveedorOSE = PrmFileCnx["ProvOSE"].ToString();
                if (PrmFileCnx.ContainsKey("CamAFE")) CamposAdicionalesFE = Convert.ToBoolean(PrmFileCnx["CamAFE"].ToString());

            }
        }
    }
}