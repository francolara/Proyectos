USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: FRANCO LARA
-- Create date: 03/05/2026
CREATE OR ALTER PROCEDURE [dbo].[Sp_Cupones_Listar]
    @NegocioId INT,
    @SedeId INT = NULL,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @Estado NVARCHAR(20) = N'vigentes',
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
            DECLARE @tmp DATE = @FechaDesde;
            SET @FechaDesde = @FechaHasta;
            SET @FechaHasta = @tmp;
        END
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);
        DECLARE @Offset INT = (@Pagina - 1) * @TamanoPagina;

        ;WITH F AS (
            SELECT
                c.Id,
                c.CodigoCupon,
                c.Nombre,
                c.TipoDescuento,
                c.ValorDescuento,
                c.CantidadMaxUsos,
                c.CantidadUsosActuales,
                (c.CantidadMaxUsos - c.CantidadUsosActuales) AS CantidadUsosDisponibles,
                c.FechaInicio,
                c.FechaFin,
                COALESCE(s.Nombre, N'Todas') AS Sede,
                COALESCE(e.Nombre, N'Todos') AS Espacio,
                c.Activo,
                CAST(CASE WHEN c.Activo = 1 AND c.FechaInicio <= @Hoy AND c.FechaFin >= @Hoy THEN 1 ELSE 0 END AS BIT) AS VigenteHoy
            FROM dbo.Cupones c
            LEFT JOIN dbo.Sedes s ON s.Id = c.SedeId
            LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = c.EspacioDeportivoId
            WHERE c.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR c.SedeId = @SedeId OR (c.SedeId IS NULL AND c.EspacioDeportivoId IS NULL))
              AND (@FechaDesde IS NULL OR c.FechaFin >= @FechaDesde)
              AND (@FechaHasta IS NULL OR c.FechaInicio <= @FechaHasta)
              AND (
                    @Estado = N'todos'
                    OR (@Estado = N'vigentes' AND c.Activo = 1 AND c.FechaInicio <= @Hoy AND c.FechaFin >= @Hoy AND c.CantidadUsosActuales < c.CantidadMaxUsos)
                    OR (@Estado = N'activos' AND c.Activo = 1)
                    OR (@Estado = N'agotados' AND c.CantidadUsosActuales >= c.CantidadMaxUsos)
                    OR (@Estado = N'vencidos' AND c.FechaFin < @Hoy)
                  )
        )
        SELECT @TotalRegistros = COUNT(1) FROM F;

        ;WITH F AS (
            SELECT
                c.Id,
                c.CodigoCupon,
                c.Nombre,
                c.TipoDescuento,
                c.ValorDescuento,
                c.CantidadMaxUsos,
                c.CantidadUsosActuales,
                (c.CantidadMaxUsos - c.CantidadUsosActuales) AS CantidadUsosDisponibles,
                c.FechaInicio,
                c.FechaFin,
                COALESCE(s.Nombre, N'Todas') AS Sede,
                COALESCE(e.Nombre, N'Todos') AS Espacio,
                c.Activo,
                CAST(CASE WHEN c.Activo = 1 AND c.FechaInicio <= @Hoy AND c.FechaFin >= @Hoy THEN 1 ELSE 0 END AS BIT) AS VigenteHoy
            FROM dbo.Cupones c
            LEFT JOIN dbo.Sedes s ON s.Id = c.SedeId
            LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = c.EspacioDeportivoId
            WHERE c.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR c.SedeId = @SedeId OR (c.SedeId IS NULL AND c.EspacioDeportivoId IS NULL))
              AND (@FechaDesde IS NULL OR c.FechaFin >= @FechaDesde)
              AND (@FechaHasta IS NULL OR c.FechaInicio <= @FechaHasta)
              AND (
                    @Estado = N'todos'
                    OR (@Estado = N'vigentes' AND c.Activo = 1 AND c.FechaInicio <= @Hoy AND c.FechaFin >= @Hoy AND c.CantidadUsosActuales < c.CantidadMaxUsos)
                    OR (@Estado = N'activos' AND c.Activo = 1)
                    OR (@Estado = N'agotados' AND c.CantidadUsosActuales >= c.CantidadMaxUsos)
                    OR (@Estado = N'vencidos' AND c.FechaFin < @Hoy)
                  )
        )
        SELECT *
        FROM F
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
