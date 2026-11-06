-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos limitados de mantenimientos, etapa 3.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCatalogos TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCliente_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCliente_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantCliente_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTienda_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTienda_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTienda_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantContactoCliente_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantContactoCliente_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantContactoCliente_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantProveedorCliente_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantProveedorCliente_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantProveedorCliente_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantUnidadMedida_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantUnidadMedida_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantUnidadMedida_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantZona_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantZona_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantZona_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantEspecialidad_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantEspecialidad_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantEspecialidad_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTipoServicio_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTipoServicio_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTipoServicio_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantServicioMaterial_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantServicioMaterial_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantServicioMaterial_Guardar TO hyt_app_executor;
GO
