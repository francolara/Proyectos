-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Concede solo operaciones web de acceso; bootstrap reservado a despliegue.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_ConsultarAcceso TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_FinalizarAcceso TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_RegistrarSalida TO hyt_app_executor;
-- No se concede CREAR_ADMIN ni RESTABLECER al rol web en esta etapa.
-- Las herramientas administrativas usan la cuenta de despliegue autorizada.
GO
