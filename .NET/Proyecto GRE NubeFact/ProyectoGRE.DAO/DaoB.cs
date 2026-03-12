using DAO;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;

namespace ProyectoGRE.DAO
{
    [Serializable]
    public class DaoB
    {
        protected SqlConnection objCn;

        public DaoB()
        {
            Conexion cn = new Conexion();
            objCn = new SqlConnection(cn.StrCon);
        }
    }
}
