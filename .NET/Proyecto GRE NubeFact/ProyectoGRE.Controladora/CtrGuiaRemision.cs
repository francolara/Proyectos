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
    public class CtrGuiaRemision
    {
        public Dto_GuiaRemision Lista_GR(DtoB dtoBase)
        {
            DaoGuiaRemision dao = new DaoGuiaRemision();
            return dao.Lista_GR(dtoBase);
        }

        public Dto_GuiaRemision Lista_Consulta_Comp_GR(DtoB dtoBase)
        {
            DaoGuiaRemision dao = new DaoGuiaRemision();
            return dao.Lista_Consulta_Comp_GR(dtoBase);
        }

        public Dto_GuiaRemision Consultar_Comp_GR(DtoRespuestaPSE_OSE_NUBEFACT _dtoENV)
        {
            CtrIntegracionPSE_OSE_NUBEFACT CtrOse = new CtrIntegracionPSE_OSE_NUBEFACT();
            CtrOse.CONSULTA_PSE_OSE_NUBEFACT(_dtoENV);
            return null;
        }


        public Dto_GuiaRemision Act_Estado_GR(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            DaoGuiaRemision dao = new DaoGuiaRemision();
            dao.Act_Estado_GR(_dtoRES);
            return null;
        }

        public Dto_GuiaRemision Consulta_Estado_GR(DtoRespuestaPSE_OSE_NUBEFACT _dtoRES)
        {
            DaoGuiaRemision dao = new DaoGuiaRemision();
            dao.Consulta_Estado_GR(_dtoRES);
            return null;
        }

        public Dto_GuiaRemision Envia_GR(DtoB dtoBase)
        {
            DaoGuiaRemision dao = new DaoGuiaRemision();

            CtrIntegracionPSE_OSE_NUBEFACT CtrOse = new CtrIntegracionPSE_OSE_NUBEFACT();
            CtrOse.Enviar_PSE_OSE_NUBEFACT(dao.Envia_GR(dtoBase), dao.Envia_GR_Det(dtoBase));

            return null;
        }

    }
}
