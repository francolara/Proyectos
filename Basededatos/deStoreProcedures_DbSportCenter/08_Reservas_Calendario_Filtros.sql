-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Filtros de reservas por rango/sede/espacio/estado para vista calendario semanal.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Listar
    @NegocioId INT,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (300)
            r.Id,
            c.NombresORazonSocial AS Cliente,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            CAST(r.Estado AS NVARCHAR(20)) AS Estado
        FROM dbo.Reservas r
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@FechaDesde IS NULL OR r.Fecha >= @FechaDesde)
          AND (@FechaHasta IS NULL OR r.Fecha <= @FechaHasta)
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND (@Estado IS NULL OR r.Estado = @Estado)
        ORDER BY r.Fecha ASC, r.HoraInicio ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
