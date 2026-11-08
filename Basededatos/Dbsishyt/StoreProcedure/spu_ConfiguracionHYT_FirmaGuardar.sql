-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Firma de presupuesto privada en R2, concurrencia y auditoría; admite ConfigHYT/Firma y claves históricas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConfiguracionHYT_FirmaGuardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Datos NVARCHAR(2000),@VersionOriginal BINARY(8)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS u INNER JOIN dbo.AspNetUserRoles AS r ON r.UserId=u.Id INNER JOIN dbo.RolPermiso AS p ON p.RoleId=r.RoleId WHERE u.Id=@UsuarioId AND u.Activo=1 AND p.CodigoPermiso=N'CONFIGURACION.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para subir la firma.',1;
        IF ISJSON(@Datos)<>1 OR @Datos IS NULL THROW 51303,N'[HYT_VALIDACION] Firma no válida.',1;
        DECLARE @Clave NVARCHAR(100)=JSON_VALUE(@Datos,N'$.ClaveAlmacenamiento'),@Tipo NVARCHAR(50)=JSON_VALUE(@Datos,N'$.TipoContenido'),@Hash NVARCHAR(100)=JSON_VALUE(@Datos,N'$.SHA256'),@Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TamanoBytes'));
        DECLARE @NombreClave NVARCHAR(100)=CASE WHEN LEFT(@Clave,16) COLLATE Latin1_General_100_BIN2=N'ConfigHYT/Firma/' THEN SUBSTRING(@Clave,17,100) ELSE @Clave END;
        IF @Clave IS NULL OR LEN(@NombreClave)<>36 OR LEFT(@NombreClave,32) COLLATE Latin1_General_100_BIN2 LIKE N'%[^a-f0-9]%' OR RIGHT(@Clave,4) NOT IN(N'.png',N'.jpg')
            OR @Tipo IS NULL OR @Tipo NOT IN(N'image/png',N'image/jpeg') OR (@Tipo=N'image/png' AND RIGHT(@Clave,4)<>N'.png') OR (@Tipo=N'image/jpeg' AND RIGHT(@Clave,4)<>N'.jpg')
            OR @Hash IS NULL OR LEN(@Hash)<>64 OR @Hash COLLATE Latin1_General_100_BIN2 LIKE N'%[^A-Fa-f0-9]%' OR @Tamano IS NULL OR @Tamano NOT BETWEEN 1 AND 2097152
            OR NULLIF(JSON_VALUE(@Datos,N'$.NombreOriginal'),N'') IS NULL
            THROW 51305,N'[HYT_VALIDACION] Adjunta una firma JPG o PNG válida de hasta 2 MB.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(2000);
        SELECT @Anterior=c.FirmaPresupuestoJson FROM dbo.ConfiguracionHYT AS c WITH(UPDLOCK,HOLDLOCK) WHERE c.Id=1;
        IF NOT EXISTS(SELECT 1 FROM dbo.ConfiguracionHYT AS c WHERE c.Id=1 AND c.VersionFila=@VersionOriginal)
            THROW 51309,N'[HYT_CONCURRENCIA] Guarda los datos de HYT o vuelve a abrir la configuración antes de subir la firma.',1;
        UPDATE dbo.ConfiguracionHYT SET FirmaPresupuestoJson=@Datos,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=1;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Datos) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfiguracionHYT',N'1',N'ACTUALIZAR_FIRMA',NULL,NULL,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
IF DATABASE_PRINCIPAL_ID(N'hyt_app_executor') IS NOT NULL GRANT EXECUTE ON dbo.spu_ConfiguracionHYT_FirmaGuardar TO hyt_app_executor;
GO
