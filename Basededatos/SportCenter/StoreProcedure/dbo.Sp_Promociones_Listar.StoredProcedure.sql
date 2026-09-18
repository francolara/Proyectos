
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 503)
-- Firma: Codex - 13/04/2026 | Agrega filtro por rango de fechas y estado (activos/inactivos/todos) con paginacion backend 20x20 y total de registros para el listado de promociones.
-- Firma: FRANCO LARA - 17/09/2026 | Los inactivos se listan por estado sin ocultarlos por su rango de vigencia.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Promociones_Listar]
    @NegocioId INT,
    @SedeId INT = NULL,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @SoloActivos BIT = 1,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @Pagina < 1 SET @Pagina = 1;
        IF @TamanoPagina < 1 SET @TamanoPagina = 20;

        IF @FechaDesde IS NOT NULL AND @FechaHasta IS NOT NULL AND @FechaHasta < @FechaDesde
        BEGIN
            DECLARE @FechaTmp DATE = @FechaDesde;
            SET @FechaDesde = @FechaHasta;
            SET @FechaHasta = @FechaTmp;
        END

        DECLARE @Offset INT = (@Pagina - 1) * @TamanoPagina;

        ;WITH PromocionesFiltradas AS
        (
            SELECT
                p.Id,
                p.Nombre,
                COALESCE(s.Nombre, N'Todas') AS Sede,
                COALESCE(e.Nombre, N'Todos') AS Espacio,
                p.FechaInicio,
                p.FechaFin,
                p.HoraInicio,
                p.HoraFin,
                p.PorcentajeDescuento,
                p.Activo
            FROM dbo.PromocionesHorario p
            LEFT JOIN dbo.Sedes s ON s.Id = p.SedeId
            LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = p.EspacioDeportivoId
            WHERE p.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR p.SedeId = @SedeId OR (p.SedeId IS NULL AND p.EspacioDeportivoId IS NULL))
              AND (@SoloActivos IS NULL OR p.Activo = @SoloActivos)
              AND (
                    @SoloActivos = 0
                    OR (
                        (@FechaDesde IS NULL OR p.FechaFin >= @FechaDesde)
                        AND (@FechaHasta IS NULL OR p.FechaInicio <= @FechaHasta)
                    )
                  )
        )
        SELECT @TotalRegistros = COUNT(1)
        FROM PromocionesFiltradas;

        ;WITH PromocionesFiltradas AS
        (
            SELECT
                p.Id,
                p.Nombre,
                COALESCE(s.Nombre, N'Todas') AS Sede,
                COALESCE(e.Nombre, N'Todos') AS Espacio,
                p.FechaInicio,
                p.FechaFin,
                p.HoraInicio,
                p.HoraFin,
                p.PorcentajeDescuento,
                p.Activo
            FROM dbo.PromocionesHorario p
            LEFT JOIN dbo.Sedes s ON s.Id = p.SedeId
            LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = p.EspacioDeportivoId
            WHERE p.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR p.SedeId = @SedeId OR (p.SedeId IS NULL AND p.EspacioDeportivoId IS NULL))
              AND (@SoloActivos IS NULL OR p.Activo = @SoloActivos)
              AND (
                    @SoloActivos = 0
                    OR (
                        (@FechaDesde IS NULL OR p.FechaFin >= @FechaDesde)
                        AND (@FechaHasta IS NULL OR p.FechaInicio <= @FechaHasta)
                    )
                  )
        )
        SELECT
            Id,
            Nombre,
            Sede,
            Espacio,
            FechaInicio,
            FechaFin,
            HoraInicio,
            HoraFin,
            PorcentajeDescuento,
            Activo
        FROM PromocionesFiltradas
        ORDER BY FechaInicio DESC, Id DESC
        OFFSET @Offset ROWS FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
