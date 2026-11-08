-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   spu_Usuario_CambiarClave; permisos, concurrencia y auditoría atómica sin secretos.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Usuario_CambiarClave
    @UsuarioId NVARCHAR(128),@SelloEsperado NVARCHAR(MAX),@HashEsperado NVARCHAR(MAX),@PasswordHash NVARCHAR(MAX),@Obligatorio BIT,@CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NULLIF(@PasswordHash,N'') IS NULL OR LEN(@PasswordHash)>1024 OR @CorrelacionId IS NULL OR @Obligatorio IS NULL THROW 51701,N'[HYT_VALIDACION] Datos de contraseña no válidos.',1;
        IF @@TRANCOUNT<>0 THROW 51701,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;SET @Propia=1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario WITH(UPDLOCK,HOLDLOCK) WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND usuario.SecurityStamp=@SelloEsperado AND usuario.PasswordHash=@HashEsperado AND usuario.DebeCambiarClave=@Obligatorio AND (usuario.LockoutEnd IS NULL OR usuario.LockoutEnd<=SYSUTCDATETIME()))
            THROW 51703,N'[HYT_CONCURRENCIA] Tu acceso cambió. Inicia sesión nuevamente.',1;
        IF @PasswordHash=@HashEsperado THROW 51701,N'[HYT_VALIDACION] La contraseña nueva debe ser diferente.',1;
        UPDATE dbo.AspNetUsers SET PasswordHash=@PasswordHash,DebeCambiarClave=0,FechaCambioClave=SYSUTCDATETIME(),SecurityStamp=CONVERT(NVARCHAR(36),NEWID()),ConcurrencyStamp=CONVERT(NVARCHAR(36),NEWID()),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=@UsuarioId;
        DECLARE @Evento NVARCHAR(60)=CASE WHEN @Obligatorio=1 THEN N'CAMBIAR_CLAVE_OBLIGATORIA' ELSE N'CAMBIAR_CLAVE' END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'AspNetUsers',@UsuarioId,@Evento,NULL,NULL,NULL,NULL,NULL;
        COMMIT;SET @Propia=0;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
