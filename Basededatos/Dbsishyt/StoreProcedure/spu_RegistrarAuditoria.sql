-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Anexa auditoria dentro de la transaccion de negocio.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RegistrarAuditoria
    @UsuarioId NVARCHAR(128),
    @CorrelacionId UNIQUEIDENTIFIER,
    @Entidad NVARCHAR(80),
    @RegistroId NVARCHAR(100),
    @Accion NVARCHAR(60),
    @EstadoAnterior NVARCHAR(40) = NULL,
    @EstadoNuevo NVARCHAR(40) = NULL,
    @FechaEvento DATETIME2(3) = NULL,
    @Motivo NVARCHAR(1000) = NULL,
    @CambiosJson NVARCHAR(MAX) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @@TRANCOUNT = 0 OR XACT_STATE() <> 1
            THROW 51000, N'La auditoria debe participar en una transaccion activa y valida.', 1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1)
            THROW 51001, N'Usuario de auditoria no valido o desactivado.', 1;
        IF @CambiosJson IS NOT NULL AND ISJSON(@CambiosJson) <> 1
            THROW 51002, N'El detalle de cambios debe ser JSON valido.', 1;

        INSERT INTO dbo.EventoAuditoria
            (CorrelacionId, Entidad, RegistroId, Accion, EstadoAnterior, EstadoNuevo,
             UsuarioId, FechaEvento, Motivo, CambiosJson)
        VALUES
            (@CorrelacionId, @Entidad, @RegistroId, @Accion, @EstadoAnterior, @EstadoNuevo,
             @UsuarioId, @FechaEvento, @Motivo, @CambiosJson);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
