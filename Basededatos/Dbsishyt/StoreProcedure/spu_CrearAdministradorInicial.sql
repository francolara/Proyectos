-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Aprovisiona primer administrador desde herramienta de despliegue.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CrearAdministradorInicial
    @UsuarioId NVARCHAR(128),
    @Usuario NVARCHAR(256),
    @NombreNormalizado NVARCHAR(256),
    @Nombre NVARCHAR(150),
    @PasswordHash NVARCHAR(MAX),
    @SecurityStamp NVARCHAR(MAX),
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0 THROW 51102, N'El aprovisionamiento administra su propia transaccion.', 1;
        IF NULLIF(@Usuario, N'') IS NULL OR NULLIF(@Nombre, N'') IS NULL OR NULLIF(@PasswordHash, N'') IS NULL OR NULLIF(@SecurityStamp, N'') IS NULL
            THROW 51103, N'Datos de aprovisionamiento incompletos.', 1;
        BEGIN TRANSACTION;
        SET @Propia = 1;
        DECLARE @BloqueoBootstrap INT;
        EXEC @BloqueoBootstrap = sys.sp_getapplock @Resource = N'HYT:BootstrapAdmin', @LockMode = N'Exclusive', @LockOwner = N'Transaction', @LockTimeout = 10000;
        IF @BloqueoBootstrap < 0 THROW 51109, N'No se pudo bloquear el aprovisionamiento inicial.', 1;
        IF EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario WITH (UPDLOCK, HOLDLOCK)
            INNER JOIN dbo.AspNetUserRoles AS asignacion WITH (UPDLOCK, HOLDLOCK) ON asignacion.UserId = usuario.Id
            WHERE asignacion.RoleId = N'HYT-ADMIN' AND usuario.Activo = 1
        )
            THROW 51104, N'Ya existe un administrador activo; use el mantenimiento autorizado.', 1;
        INSERT INTO dbo.AspNetUsers (Id, UserName, NormalizedUserName, Nombres, PasswordHash,
            SecurityStamp, ConcurrencyStamp, Activo, LockoutEnabled)
        VALUES (@UsuarioId, @Usuario, @NombreNormalizado, @Nombre, @PasswordHash,
            @SecurityStamp, CONVERT(NVARCHAR(36), NEWID()), 1, 1);
        INSERT INTO dbo.AspNetUserRoles (UserId, RoleId) VALUES (@UsuarioId, N'HYT-ADMIN');
        DECLARE @Detalle NVARCHAR(MAX) = (SELECT ORIGINAL_LOGIN() AS LoginDespliegue FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId, Entidad, RegistroId, Accion, UsuarioId, CambiosJson)
        VALUES (@CorrelacionId, N'AspNetUsers', @UsuarioId, N'ADMINISTRADOR_INICIAL', @UsuarioId, @Detalle);
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
