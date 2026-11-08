-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/11/2026
-- Description:   Autoriza la ejecución de las operaciones administrativas públicas de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_Seguridad_Catalogos TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuarios_Accion TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Perfiles_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Perfiles_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Perfiles_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Usuario_CambiarClave TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Auditoria_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Auditoria_Obtener TO hyt_app_executor;
GO
