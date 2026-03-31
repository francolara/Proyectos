-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Incluye eventos de no atencion de sede en calendario (dias no laborables y fechas inhabilitadas) y devuelve todos los estados de reserva.
-- Firma:         Codex - 27/03/2026
-- Firma:         Codex - 28/03/2026 | Ajuste de estados en calendario (incluye canceladas/no show), correccion de Id NO_ATENCION, franjas fuera de horario y color bloqueado unificado (#64748b).
-- Firma:         Codex - 30/03/2026 | Calendario backend-driven: agrega Motivo, EstadoCodigo y EstadoTexto para eliminar fallback en C#/FrontEnd.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_CalendarioEventos
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH Fechas AS
        (
            SELECT @FechaDesde AS Fecha
            UNION ALL
            SELECT DATEADD(DAY, 1, Fecha) FROM Fechas WHERE Fecha < @FechaHasta
        )
        SELECT
            r.Id,
            CAST(N'RESERVA' AS NVARCHAR(20)) AS TipoEvento,
            CONCAT(e.Nombre, N' - ', c.NombresORazonSocial) AS Titulo,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Estado,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'#f59f00'
                    WHEN 2 THEN N'#2f9e44'
                    WHEN 3 THEN N'#1971c2'
                    WHEN 4 THEN N'#495057'
                    WHEN 5 THEN N'#c92a2a'
                    WHEN 6 THEN N'#212529'
                    ELSE N'#6c757d'
                END
                AS NVARCHAR(20)
            ) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            CAST(NULL AS NVARCHAR(200)) AS Motivo,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'PENDIENTE'
                    WHEN 2 THEN N'CONFIRMADA'
                    WHEN 3 THEN N'EN_USO'
                    WHEN 4 THEN N'FINALIZADA'
                    WHEN 5 THEN N'CANCELADA'
                    WHEN 6 THEN N'NO_SHOW'
                    ELSE N'RESERVADA'
                END
                AS NVARCHAR(40)
            ) AS EstadoCodigo,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'Pendiente'
                    WHEN 2 THEN N'Confirmada'
                    WHEN 3 THEN N'En uso'
                    WHEN 4 THEN N'Finalizada'
                    WHEN 5 THEN N'Cancelada'
                    WHEN 6 THEN N'No show'
                    ELSE N'Reservada'
                END
                AS NVARCHAR(80)
            ) AS EstadoTexto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
          AND r.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND (@Estado IS NULL OR r.Estado = @Estado)

        UNION ALL

        SELECT
            b.Id,
            CAST(N'BLOQUEO' AS NVARCHAR(20)) AS TipoEvento,
            CONCAT(N'Bloqueado: ', b.Motivo) AS Titulo,
            b.Fecha,
            b.HoraInicio,
            b.HoraFin,
            NULL AS Estado,
            CAST(N'#64748b' AS NVARCHAR(20)) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            b.Motivo AS Motivo,
            CAST(N'BLOQUEADO' AS NVARCHAR(40)) AS EstadoCodigo,
            CAST(N'Bloqueado' AS NVARCHAR(80)) AS EstadoTexto
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND b.Activo = 1
          AND b.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)

        UNION ALL

        SELECT
            (
                110000000
                + (DATEDIFF(DAY, '2020-01-01', sfi.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (fecha inhabilitada)' AS NVARCHAR(200)),
            sfi.Fecha,
            CAST('00:00' AS TIME),
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fecha inhabilitada)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM dbo.SedeFechasInhabilitadas sfi
        INNER JOIN dbo.Sedes s ON s.Id = sfi.SedeId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND sfi.Activo = 1
          AND sfi.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)

        UNION ALL

        SELECT
            (
                120000000
                + (DATEDIFF(DAY, '2020-01-01', f.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (dia no laborable)' AS NVARCHAR(200)),
            f.Fecha,
            CAST('00:00' AS TIME),
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (dia no laborable)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
              END = 0

        UNION ALL

        SELECT
            (
                130000000
                + (DATEDIFF(DAY, '2020-01-01', f.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            f.Fecha,
            CAST('00:00' AS TIME),
            COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
              END = 1
          AND COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) > CAST('00:00' AS TIME)

        UNION ALL

        SELECT
            (
                140000000
                + (DATEDIFF(DAY, '2020-01-01', f.Fecha) * 10000)
                + (e.Id % 10000)
            ),
            CAST(N'NO_ATENCION' AS NVARCHAR(20)),
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            f.Fecha,
            COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)),
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN COALESCE(sha.AtiendeLunes, 1)
                WHEN 2 THEN COALESCE(sha.AtiendeMartes, 1)
                WHEN 3 THEN COALESCE(sha.AtiendeMiercoles, 1)
                WHEN 4 THEN COALESCE(sha.AtiendeJueves, 1)
                WHEN 5 THEN COALESCE(sha.AtiendeViernes, 1)
                WHEN 6 THEN COALESCE(sha.AtiendeSabado, 1)
                WHEN 7 THEN COALESCE(sha.AtiendeDomingo, 1)
              END = 1
          AND COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) < CAST('23:59' AS TIME)

        ORDER BY Fecha, HoraInicio
        OPTION (MAXRECURSION 400);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
