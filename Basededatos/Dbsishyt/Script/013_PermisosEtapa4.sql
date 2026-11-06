-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos limitados de mantenimientos, etapa 4.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Agrega consulta de precios por material con permisos existentes.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifario_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifario_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifario_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioPartida_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioPartida_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioPartida_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MaterialTarifas_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MaterialTarifa_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Publicar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Clonar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTasaIGV_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTasaIGV_Agregar TO hyt_app_executor;
GO
