-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos individuales de presupuestos, etapa 5.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_AyudaTarifario TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_DatosIniciales TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_ListarImpresiones TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_ObtenerImpresion TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_RegistrarImpresion TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ConfiguracionHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ConfiguracionHYT_Obtener TO hyt_app_executor;
GO
