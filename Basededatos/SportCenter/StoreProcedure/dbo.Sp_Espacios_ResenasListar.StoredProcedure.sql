GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/06/2026
-- Description:   Lista las resenas registradas para un espacio deportivo del negocio con estado visible y respuesta administrativa.
-- =============================================
-- Firma:         FRANCO LARA - 11/06/2026 | Permite al administrador del negocio revisar, responder y moderar resenas desde el listado de espacios.
-- Firma:         FRANCO LARA - 11/06/2026 | Aplica paginacion SQL y devuelve KPIs globales para la gestion administrativa del espacio.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_ResenasListar]
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @Pagina INT = 1,
    @TamanoPagina INT = 4,
    @TotalRegistros INT OUTPUT,
    @TotalVisibles INT OUTPUT,
    @TotalOcultas INT OUTPUT,
    @TotalRespondidas INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 4 ELSE @TamanoPagina END;

        SELECT
            @TotalRegistros = COUNT(1),
            @TotalVisibles = SUM(CASE WHEN rr.Activo = 1 THEN 1 ELSE 0 END),
            @TotalOcultas = SUM(CASE WHEN rr.Activo = 0 THEN 1 ELSE 0 END),
            @TotalRespondidas = SUM(CASE WHEN NULLIF(LTRIM(RTRIM(rr.Respuesta)), N'') IS NOT NULL THEN 1 ELSE 0 END)
        FROM dbo.ReservasUsuariosPublicosResenas rr
        INNER JOIN dbo.Reservas r ON r.Id = rr.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND e.Id = @EspacioDeportivoId;

        SELECT
            @TotalRegistros = ISNULL(@TotalRegistros, 0),
            @TotalVisibles = ISNULL(@TotalVisibles, 0),
            @TotalOcultas = ISNULL(@TotalOcultas, 0),
            @TotalRespondidas = ISNULL(@TotalRespondidas, 0);

        SELECT
            rr.Id,
            rr.ReservaId,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS EspacioNombre,
            s.Id AS SedeId,
            s.Nombre AS SedeNombre,
            rr.AliasPublico,
            rr.Comentario,
            rr.Respuesta,
            rr.Activo,
            rr.FechaCreacion,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin
        FROM dbo.ReservasUsuariosPublicosResenas rr
        INNER JOIN dbo.Reservas r ON r.Id = rr.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND e.Id = @EspacioDeportivoId
        ORDER BY rr.FechaCreacion DESC, rr.Id DESC
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
