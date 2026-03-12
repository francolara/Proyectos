//using Controladora.Eventos;
using DAO;
//using DTO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using ProyectoGRE.DTO;
using ProyectoGRE.DAO;
using Controladora.IntegracionPSE_OSE_SUNAT;

namespace ProyectoGRE.Controladora
{
    public class CtrDocFacturacion
    {
        public Dto_GuiaRemision Lista_DocFac(DtoB dtoBase)
        {
            DaoDocFacturacion dao = new DaoDocFacturacion();
            return dao.Lista_DocFac(dtoBase);
        }

        public Dto_GuiaRemision Lista_Consulta_Comp_DocFac(DtoB dtoBase)
        {
            DaoDocFacturacion dao = new DaoDocFacturacion();
            return dao.Lista_Consulta_Comp_DocFac(dtoBase);
        }

        public Dto_GuiaRemision Consultar_Comp_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            CtrIntegracionPSE_OSE_NUBEFACT CtrOse = new CtrIntegracionPSE_OSE_NUBEFACT();
            CtrOse.CONSULTA_PSE_OSE_NUBEFACT_DocFac(_dtoENV);
            return null;
        }


        public Dto_GuiaRemision Act_Estado_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            DaoDocFacturacion dao = new DaoDocFacturacion();
            dao.Act_Estado_DocFac(_dtoRES);
            return null;
        }

        public Dto_GuiaRemision Consulta_Estado_DocFac(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            DaoDocFacturacion dao = new DaoDocFacturacion();
            dao.Consulta_Estado_DocFac(_dtoRES);
            return null;
        }

        public Dto_GuiaRemision Envia_DocFac(DtoB dtoBase)
        {
            DaoDocFacturacion dao = new DaoDocFacturacion();

            CtrIntegracionPSE_OSE_NUBEFACT CtrOse = new CtrIntegracionPSE_OSE_NUBEFACT();
            CtrOse.Enviar_PSE_OSE_NUBEFACT_DocFac(dao.Envia_DocFac(dtoBase) , dao.Envia_DocFac_Det(dtoBase),  dao.Envia_DocFac_Det_GR(dtoBase));

            return null;
        }

    }
}
