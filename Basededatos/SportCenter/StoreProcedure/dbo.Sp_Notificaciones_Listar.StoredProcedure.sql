USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Lista notificaciones recientes para campanita admin por negocio.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Notificaciones_Listar]
    @NegocioId INT,
    @Top INT = 15
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Top IS NULL OR @Top <= 0 SET @Top = 15;
        IF @Top > 100 SET @Top = 100;

        SELECT TOP (@Top)
            n.Id,
            n.Tipo,
            n.Titulo,
            n.Mensaje,
            n.UrlDestino,
            n.FechaRegistroUtc
        FROM dbo.NegocioNotificaciones n
        WHERE n.NegocioId = @NegocioId
          AND n.Leida = 0
        ORDER BY n.FechaRegistroUtc DESC, n.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
