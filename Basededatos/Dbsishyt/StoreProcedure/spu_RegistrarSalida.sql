-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Registra cierre de sesion incluso si la cuenta fue desactivada.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RegistrarSalida
    @UsuarioId NVARCHAR(128),
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO dbo.EventoAuditoria (CorrelacionId, Entidad, RegistroId, Accion, UsuarioId)
        SELECT @CorrelacionId, N'AspNetUsers', usuario.Id, N'CIERRE_SESION', usuario.Id
        FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
