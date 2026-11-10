-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Habilita los procedimientos de cargos con permisos de Trabajadores.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CargoTrabajadorHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CargoTrabajadorHYT_Guardar TO hyt_app_executor;
GO
