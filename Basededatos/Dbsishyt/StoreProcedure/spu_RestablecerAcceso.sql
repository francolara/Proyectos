-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Restablecimiento administrativo con hash, sello nuevo y auditoria.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RestablecerAcceso
    @ActorId NVARCHAR(128),
    @UsuarioNormalizado NVARCHAR(256),
    @PasswordHash NVARCHAR(MAX),
    @SecurityStamp NVARCHAR(MAX),
    @Motivo NVARCHAR(1000),
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0 THROW 51105, N'El restablecimiento administra su propia transaccion.', 1;
        IF NULLIF(LTRIM(RTRIM(@Motivo)), N'') IS NULL OR NULLIF(@PasswordHash, N'') IS NULL OR NULLIF(@SecurityStamp, N'') IS NULL
            THROW 51106, N'Indique hash, sello y motivo.', 1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
            WHERE usuario.Id = @ActorId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'USUARIOS.EDITAR'
        )
            THROW 51107, N'Actor sin permiso para restablecer acceso.', 1;
        BEGIN TRANSACTION;
        SET @Propia = 1;
        DECLARE @UsuarioId NVARCHAR(128);
        SELECT @UsuarioId = usuario.Id FROM dbo.AspNetUsers AS usuario WITH (UPDLOCK, HOLDLOCK)
        WHERE usuario.NormalizedUserName = @UsuarioNormalizado;
        IF @UsuarioId IS NULL THROW 51108, N'Usuario no encontrado.', 1;
        UPDATE dbo.AspNetUsers SET PasswordHash = @PasswordHash, SecurityStamp = @SecurityStamp,
            ConcurrencyStamp = CONVERT(NVARCHAR(36), NEWID()), AccessFailedCount = 0, LockoutEnd = NULL,
            FechaActualizacion = SYSUTCDATETIME() WHERE Id = @UsuarioId;
        DECLARE @Detalle NVARCHAR(MAX) = (SELECT @UsuarioId AS UsuarioAfectado, ORIGINAL_LOGIN() AS LoginDespliegue FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @ActorId, @CorrelacionId, N'AspNetUsers', @UsuarioId,
            N'RESTABLECER_ACCESO', NULL, NULL, NULL, @Motivo, @Detalle;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
