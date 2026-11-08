-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda datos únicos HYT con concurrencia y auditoría, sin reescribir revisiones.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo en direcciones y consideraciones predeterminadas por presupuesto.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Consulta y guardado de ubigeo HYT con concurrencia y auditoría. Subtítulo de cotización editable con auditoría.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConfiguracionHYT_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Datos NVARCHAR(MAX),@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIGURACION.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para editar la configuración.',1;
        IF ISJSON(@Datos)<>1 THROW 51303,N'[HYT_VALIDACION] Datos no válidos.',1;
        DECLARE @Ruc NVARCHAR(20)=JSON_VALUE(@Datos,N'$.Ruc'),@Razon NVARCHAR(200)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RazonSocial'))),N'');
        IF @Ruc IS NULL OR LEN(@Ruc)<>11 OR @Ruc LIKE N'%[^0-9]%' OR @Razon IS NULL
            THROW 51305,N'[HYT_VALIDACION] Completa RUC de once dígitos y razón social.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo INNER JOIN (VALUES (N'Ruc',11),(N'RazonSocial',200),(N'Direccion',300),(N'CodigoUbigeo',6),(N'Telefono',50),(N'Correo',256),(N'NombreFirmante',150),(N'CargoFirmante',150),(N'ConsideracionesPresupuesto',20000),(N'SubtituloCotizacion',2000)) AS limite(Nombre,Largo)
            ON limite.Nombre=campo.[key] COLLATE DATABASE_DEFAULT WHERE LEN(campo.value)>limite.Largo)
            THROW 51306,N'[HYT_VALIDACION] Un campo supera su longitud permitida.',1;
        DECLARE @CodigoUbigeo CHAR(6)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51305,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@Existe BIT=0;
        SELECT @Existe=1,@Anterior=(SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.CodigoUbigeo,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.SubtituloCotizacion,empresa.ConsideracionesPresupuesto FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ConfiguracionHYT AS empresa WITH (UPDLOCK,HOLDLOCK) WHERE empresa.Id=1;
        IF @Existe=1 AND NOT EXISTS (SELECT 1 FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1 AND empresa.VersionFila=@VersionOriginal)
            THROW 51309,N'[HYT_CONCURRENCIA] La configuración cambió. Vuelve a consultarla.',1;
        IF @Existe=0 AND @VersionOriginal IS NOT NULL THROW 51309,N'[HYT_CONCURRENCIA] La configuración ya no existe. Vuelve a consultarla.',1;
        IF @Existe=0 INSERT INTO dbo.ConfiguracionHYT (Id,Ruc,RazonSocial,UsuarioRegistro) VALUES (1,@Ruc,@Razon,@UsuarioId);
        UPDATE dbo.ConfiguracionHYT SET SubtituloCotizacion=CASE WHEN EXISTS(SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'SubtituloCotizacion') THEN JSON_VALUE(@Datos,N'$.SubtituloCotizacion') ELSE SubtituloCotizacion END,Ruc=@Ruc,RazonSocial=@Razon,Direccion=JSON_VALUE(@Datos,N'$.Direccion'),CodigoUbigeo=@CodigoUbigeo,Telefono=JSON_VALUE(@Datos,N'$.Telefono'),Correo=JSON_VALUE(@Datos,N'$.Correo'),
            NombreFirmante=JSON_VALUE(@Datos,N'$.NombreFirmante'),CargoFirmante=JSON_VALUE(@Datos,N'$.CargoFirmante'),ConsideracionesPresupuesto=(SELECT dato.value FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'ConsideracionesPresupuesto'),
            FechaActualizacion=CASE WHEN @Existe=1 THEN SYSUTCDATETIME() ELSE NULL END,UsuarioActualizacion=CASE WHEN @Existe=1 THEN @UsuarioId ELSE NULL END WHERE Id=1;
        SELECT @Nuevo=(SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.CodigoUbigeo,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.SubtituloCotizacion,empresa.ConsideracionesPresupuesto FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),@Accion NVARCHAR(60)=CASE WHEN @Existe=0 THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfiguracionHYT',N'1',@Accion,NULL,NULL,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
