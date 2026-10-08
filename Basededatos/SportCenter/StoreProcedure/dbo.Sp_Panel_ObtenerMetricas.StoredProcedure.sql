
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 397)
-- Firma: Codex - 13/04/2026 | Ajusta ocupacion del dashboard para calcular horas disponibles netas (horario sede menos bloqueos activos del dia) y excluir periodos inhabilitados.
-- Firma: FRANCO LARA - 06/10/2026 | Filtra reservas e importes por CodigoMoneda canonico del dashboard.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Panel_ObtenerMetricas]
    @NegocioId INT,
    @Fecha DATE,
    @SedeId INT = NULL,
    @CodigoMoneda NVARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @TotalSedes INT = 0, @TotalEspacios INT = 0, @ReservasHoy INT = 0;
        DECLARE @IngresosHoy DECIMAL(12,2) = 0, @OcupacionHoyPct DECIMAL(5,2) = 0;
        DECLARE @NoShowMes INT = 0, @TicketPromedioMes DECIMAL(12,2) = 0;

        SELECT @TotalSedes = COUNT(1)
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId);

        SELECT @TotalEspacios = COUNT(1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND e.Estado = 1;

        SELECT @ReservasHoy = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.CodigoMoneda = @CodigoMoneda
          AND r.Fecha = @Fecha;

        SELECT @IngresosHoy = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND p.CodigoMoneda = @CodigoMoneda
          AND CAST(p.FechaPago AS DATE) = @Fecha;

        DECLARE @TotalMinDisponibles INT = 0;
        DECLARE @TotalMinReservados INT = 0;

        ;WITH EspaciosBase AS
        (
            SELECT
                e.Id AS EspacioDeportivoId,
                s.Id AS SedeId,
                COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) AS HoraApertura,
                COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) AS HoraCierre,
                CASE (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1
                    WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                    WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                    WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                    WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                    WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                    WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                    WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
                    ELSE 0
                END AS DiaHabilitado,
                CASE
                    WHEN EXISTS (
                        SELECT 1
                        FROM dbo.SedeFechasInhabilitadas sfi
                        WHERE sfi.SedeId = s.Id
                          AND sfi.Fecha = @Fecha
                          AND sfi.Activo = 1
                    ) THEN 1
                    ELSE 0
                END AS FechaInhabilitada
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND e.Estado = 1
        ),
        DisponibilidadEspacio AS
        (
            SELECT
                eb.EspacioDeportivoId,
                CASE
                    WHEN eb.FechaInhabilitada = 1 OR eb.DiaHabilitado = 0 THEN 0
                    ELSE
                        CASE
                            WHEN DATEDIFF(MINUTE, eb.HoraApertura, eb.HoraCierre) > 0
                                THEN DATEDIFF(MINUTE, eb.HoraApertura, eb.HoraCierre)
                            ELSE 0
                        END
                END AS MinutosProgramados,
                CASE
                    WHEN eb.FechaInhabilitada = 1 OR eb.DiaHabilitado = 0 THEN 0
                    ELSE COALESCE(bh.MinutosBloqueados, 0)
                END AS MinutosBloqueados
            FROM EspaciosBase eb
            OUTER APPLY
            (
                SELECT SUM(
                    DATEDIFF(
                        MINUTE,
                        limites.InicioClip,
                        limites.FinClip
                    )
                ) AS MinutosBloqueados
                FROM dbo.BloqueosHorario b
                CROSS APPLY
                (
                    SELECT
                        CASE WHEN b.HoraInicio > eb.HoraApertura THEN b.HoraInicio ELSE eb.HoraApertura END AS InicioClip,
                        CASE WHEN b.HoraFin < eb.HoraCierre THEN b.HoraFin ELSE eb.HoraCierre END AS FinClip
                ) limites
                WHERE b.EspacioDeportivoId = eb.EspacioDeportivoId
                  AND b.Fecha = @Fecha
                  AND b.Activo = 1
                  AND b.HoraFin > eb.HoraApertura
                  AND b.HoraInicio < eb.HoraCierre
                  AND limites.FinClip > limites.InicioClip
            ) bh
        )
        SELECT
            @TotalMinDisponibles = COALESCE(SUM(
                CASE
                    WHEN de.MinutosProgramados <= de.MinutosBloqueados THEN 0
                    ELSE de.MinutosProgramados - de.MinutosBloqueados
                END
            ), 0)
        FROM DisponibilidadEspacio de;

        SELECT
            @TotalMinReservados = COALESCE(SUM(DATEDIFF(MINUTE, r.HoraInicio, r.HoraFin)), 0)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.CodigoMoneda = @CodigoMoneda
          AND e.Estado = 1
          AND r.Fecha = @Fecha
          AND r.Estado NOT IN (5, 6);

        IF @TotalMinDisponibles > 0
        BEGIN
            DECLARE @OcupacionCalc DECIMAL(9,4) = CAST((@TotalMinReservados * 100.0) / @TotalMinDisponibles AS DECIMAL(9,4));
            SET @OcupacionHoyPct = CAST(CASE WHEN @OcupacionCalc > 100 THEN 100 ELSE @OcupacionCalc END AS DECIMAL(5,2));
        END;

        SELECT @NoShowMes = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.CodigoMoneda = @CodigoMoneda
          AND r.Estado = 6
          AND YEAR(r.Fecha) = YEAR(@Fecha)
          AND MONTH(r.Fecha) = MONTH(@Fecha);

        DECLARE @TotalCobradoMes DECIMAL(12,2) = 0, @ReservasPagadasMes INT = 0;

        SELECT @TotalCobradoMes = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND p.CodigoMoneda = @CodigoMoneda
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        SELECT @ReservasPagadasMes = COUNT(DISTINCT p.ReservaId)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND p.CodigoMoneda = @CodigoMoneda
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        IF @ReservasPagadasMes > 0
            SET @TicketPromedioMes = CAST(@TotalCobradoMes / @ReservasPagadasMes AS DECIMAL(12,2));

        SELECT
            @TotalSedes AS TotalSedes,
            @TotalEspacios AS TotalEspacios,
            @ReservasHoy AS ReservasHoy,
            @IngresosHoy AS IngresosHoy,
            @OcupacionHoyPct AS OcupacionHoyPct,
            @NoShowMes AS NoShowMes,
            @TicketPromedioMes AS TicketPromedioMes;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
