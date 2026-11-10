-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Concede al rol de aplicacion la ejecucion de los procedimientos nuevos de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CeseTrabajadorHYT_Reingresar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_EstimadoPrevisionalHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_EntidadPrevisionalHYT_GuardarTasa TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PlanPagoHYT_ConfirmarPago TO hyt_app_executor;
GO
