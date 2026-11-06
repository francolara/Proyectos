-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos individuales de seguimiento documental, etapa 6.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Seguimiento_ObtenerArchivo TO hyt_app_executor;
GO
