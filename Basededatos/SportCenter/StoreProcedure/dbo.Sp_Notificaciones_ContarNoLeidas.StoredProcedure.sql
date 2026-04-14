USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Cuenta notificaciones no leidas por negocio para badge de campanita.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Notificaciones_ContarNoLeidas]
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT COUNT(1) AS TotalNoLeidas
        FROM dbo.NegocioNotificaciones n
        WHERE n.NegocioId = @NegocioId
          AND n.Leida = 0;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
