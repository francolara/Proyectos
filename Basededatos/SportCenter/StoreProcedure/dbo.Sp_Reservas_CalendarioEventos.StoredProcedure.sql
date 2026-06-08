
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_CalendarioEventos]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 34_Clientes_NombreEquipo_Reservas.sql (linea 316)
-- Firma: Codex - 05/04/2026 | Estados de reserva: retiro de En uso, Finalizada renombrada a Pagada y No Show a No Asistio (normalizacion de salida de calendario).
-- Firma: Codex - 08/04/2026 | Agrega TotalReserva en la salida del calendario para pintar precio en tarjetas de reservas sin incluir nombre de espacio.
-- Firma: Codex - 13/04/2026 | Calendario excluye reservas canceladas por defecto (Estado 5) para liberar horario; si se filtra Estado=5 se siguen consultando canceladas.
-- Firma: FRANCO LARA - 26/05/2026 | Usa horario de espacio cuando ConfigurarHorarioPorEspacio=1; si no, mantiene horario de sede.
-- Firma: FRANCO LARA - 06/06/2026 | Cuando se filtra un espacio, incluye reservas y bloqueos manuales de espacios compartidos como eventos bloqueantes no editables.
-- Firma: FRANCO LARA - 08/06/2026 | Distingue bloqueo directo y espacios compuestos para evitar sobrebloqueos por propagacion en cadena.
CREATE  OR ALTER PROCEDURE [dbo].[Sp_Reservas_CalendarioEventos]
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
        DECLARE @EspaciosFiltrados TABLE
        (
            EspacioDeportivoId INT NOT NULL PRIMARY KEY
        );

        IF @EspacioDeportivoId IS NOT NULL
        BEGIN
            INSERT INTO @EspaciosFiltrados (EspacioDeportivoId)
            VALUES (@EspacioDeportivoId);

            INSERT INTO @EspaciosFiltrados (EspacioDeportivoId)
            SELECT DISTINCT ec.EspacioRelacionadoId
            FROM dbo.EspaciosDeportivosCompartidos ec
            INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioRelacionadoId
            WHERE ec.EspacioDeportivoId = @EspacioDeportivoId
              AND ec.Activo = 1
              AND ec.TipoRelacion = N'DIRECTO'
              AND er.Estado = 1
              AND NOT EXISTS (SELECT 1 FROM @EspaciosFiltrados ef WHERE ef.EspacioDeportivoId = ec.EspacioRelacionadoId);

            INSERT INTO @EspaciosFiltrados (EspacioDeportivoId)
            SELECT DISTINCT ec.EspacioRelacionadoId
            FROM dbo.EspaciosDeportivosCompartidos ec
            INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioRelacionadoId
            WHERE ec.EspacioDeportivoId = @EspacioDeportivoId
              AND ec.Activo = 1
              AND ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
              AND er.Estado = 1
              AND NOT EXISTS (SELECT 1 FROM @EspaciosFiltrados ef WHERE ef.EspacioDeportivoId = ec.EspacioRelacionadoId);

            INSERT INTO @EspaciosFiltrados (EspacioDeportivoId)
            SELECT DISTINCT ec.EspacioDeportivoId
            FROM dbo.EspaciosDeportivosCompartidos ec
            INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioDeportivoId
            WHERE ec.EspacioRelacionadoId = @EspacioDeportivoId
              AND ec.Activo = 1
              AND ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
              AND er.Estado = 1
              AND NOT EXISTS (SELECT 1 FROM @EspaciosFiltrados ef WHERE ef.EspacioDeportivoId = ec.EspacioDeportivoId);

            INSERT INTO @EspaciosFiltrados (EspacioDeportivoId)
            SELECT DISTINCT ed.EspacioRelacionadoId
            FROM dbo.EspaciosDeportivosCompartidos ecComp
            INNER JOIN dbo.EspaciosDeportivosCompartidos ed
                ON ed.EspacioDeportivoId = ecComp.EspacioRelacionadoId
               AND ed.Activo = 1
               AND ed.TipoRelacion = N'DIRECTO'
            INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ed.EspacioRelacionadoId
            WHERE ecComp.EspacioDeportivoId = @EspacioDeportivoId
              AND ecComp.Activo = 1
              AND ecComp.TipoRelacion = N'COMPUESTO_COMPONENTE'
              AND er.Estado = 1
              AND NOT EXISTS (SELECT 1 FROM @EspaciosFiltrados ef WHERE ef.EspacioDeportivoId = ed.EspacioRelacionadoId);

            INSERT INTO @EspaciosFiltrados (EspacioDeportivoId)
            SELECT DISTINCT ep.EspacioDeportivoId
            FROM dbo.EspaciosDeportivosCompartidos edActual
            INNER JOIN dbo.EspaciosDeportivosCompartidos ep
                ON ep.EspacioRelacionadoId = edActual.EspacioRelacionadoId
               AND ep.Activo = 1
               AND ep.TipoRelacion = N'COMPUESTO_COMPONENTE'
            INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ep.EspacioDeportivoId
            WHERE edActual.EspacioDeportivoId = @EspacioDeportivoId
              AND edActual.Activo = 1
              AND edActual.TipoRelacion = N'DIRECTO'
              AND er.Estado = 1
              AND NOT EXISTS (SELECT 1 FROM @EspaciosFiltrados ef WHERE ef.EspacioDeportivoId = ep.EspacioDeportivoId);
        END

        ;WITH Fechas AS
        (
            SELECT @FechaDesde AS Fecha
            UNION ALL
            SELECT DATEADD(DAY, 1, Fecha) FROM Fechas WHERE Fecha < @FechaHasta
        )
        SELECT
            r.Id,
            CAST(CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'RESERVA_COMPARTIDA' ELSE N'RESERVA' END AS NVARCHAR(20)) AS TipoEvento,
            CAST(
                CASE
                    WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId
                        THEN CONCAT(N'Bloqueado por reserva en ', e.Nombre)
                    WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL
                        THEN CONCAT(e.Nombre, N' - ', c.NombresORazonSocial)
                    ELSE CONCAT(e.Nombre, N' - ', LTRIM(RTRIM(c.NombreEquipo)), N' (', c.NombresORazonSocial, N')')
                END
                AS NVARCHAR(300)
            ) AS Titulo,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Estado,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#2563eb' ELSE N'#f59f00' END
                    WHEN 2 THEN CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#2563eb' ELSE N'#2f9e44' END
                    WHEN 3 THEN CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#2563eb' ELSE N'#495057' END
                    WHEN 4 THEN CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#2563eb' ELSE N'#495057' END
                    WHEN 5 THEN CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#2563eb' ELSE N'#c92a2a' END
                    WHEN 6 THEN CASE WHEN @EspacioDeportivoId IS NOT NULL AND r.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#2563eb' ELSE N'#212529' END
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
                    WHEN 3 THEN N'PAGADA'
                    WHEN 4 THEN N'PAGADA'
                    WHEN 5 THEN N'CANCELADA'
                    WHEN 6 THEN N'NO_ASISTIO'
                    ELSE N'RESERVADA'
                END
                AS NVARCHAR(40)
            ) AS EstadoCodigo,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'Pendiente'
                    WHEN 2 THEN N'Confirmada'
                    WHEN 3 THEN N'Pagada'
                    WHEN 4 THEN N'Pagada'
                    WHEN 5 THEN N'Cancelada'
                    WHEN 6 THEN N'No Asistio'
                    ELSE N'Reservada'
                END
                AS NVARCHAR(80)
            ) AS EstadoTexto,
            CAST(ISNULL(r.Total, 0) AS DECIMAL(10,2)) AS TotalReserva
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
          AND r.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id IN (SELECT EspacioDeportivoId FROM @EspaciosFiltrados))
          AND
          (
              (@Estado IS NULL AND r.Estado <> 5)
              OR (@Estado = 4 AND r.Estado IN (3, 4))
              OR (@Estado IS NOT NULL AND @Estado <> 4 AND r.Estado = @Estado)
          )

        UNION ALL

        SELECT
            b.Id,
            CAST(CASE WHEN @EspacioDeportivoId IS NOT NULL AND b.EspacioDeportivoId <> @EspacioDeportivoId THEN N'BLOQUEO_COMPARTIDO' ELSE N'BLOQUEO' END AS NVARCHAR(20)) AS TipoEvento,
            CAST(
                CASE
                    WHEN @EspacioDeportivoId IS NOT NULL AND b.EspacioDeportivoId <> @EspacioDeportivoId
                        THEN CONCAT(N'Bloqueado por espacio compartido: ', e.Nombre)
                    ELSE CONCAT(N'Bloqueado: ', b.Motivo)
                END
                AS NVARCHAR(300)
            ) AS Titulo,
            b.Fecha,
            b.HoraInicio,
            b.HoraFin,
            NULL AS Estado,
            CAST(CASE WHEN @EspacioDeportivoId IS NOT NULL AND b.EspacioDeportivoId <> @EspacioDeportivoId THEN N'#0f766e' ELSE N'#64748b' END AS NVARCHAR(20)) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede,
            CAST(
                CASE
                    WHEN @EspacioDeportivoId IS NOT NULL AND b.EspacioDeportivoId <> @EspacioDeportivoId
                        THEN CONCAT(b.Motivo, N' | Origen: ', e.Nombre)
                    ELSE b.Motivo
                END
                AS NVARCHAR(200)
            ) AS Motivo,
            CAST(CASE WHEN @EspacioDeportivoId IS NOT NULL AND b.EspacioDeportivoId <> @EspacioDeportivoId THEN N'BLOQUEO_COMPARTIDO' ELSE N'BLOQUEADO' END AS NVARCHAR(40)) AS EstadoCodigo,
            CAST(CASE WHEN @EspacioDeportivoId IS NOT NULL AND b.EspacioDeportivoId <> @EspacioDeportivoId THEN N'Bloqueado por espacio compartido' ELSE N'Bloqueado' END AS NVARCHAR(80)) AS EstadoTexto,
            CAST(0 AS DECIMAL(10,2)) AS TotalReserva
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND b.Activo = 1
          AND b.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id IN (SELECT EspacioDeportivoId FROM @EspaciosFiltrados))

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
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80)),
            CAST(0 AS DECIMAL(10,2))
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
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80)),
            CAST(0 AS DECIMAL(10,2))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND CASE ((DATEDIFF(DAY, '19000101', f.Fecha) % 7) + 1)
                WHEN 1 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeLunes, 1) ELSE COALESCE(sha.AtiendeLunes, 1) END
                WHEN 2 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeMartes, 1) ELSE COALESCE(sha.AtiendeMartes, 1) END
                WHEN 3 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeMiercoles, 1) ELSE COALESCE(sha.AtiendeMiercoles, 1) END
                WHEN 4 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeJueves, 1) ELSE COALESCE(sha.AtiendeJueves, 1) END
                WHEN 5 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeViernes, 1) ELSE COALESCE(sha.AtiendeViernes, 1) END
                WHEN 6 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeSabado, 1) ELSE COALESCE(sha.AtiendeSabado, 1) END
                WHEN 7 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeDomingo, 1) ELSE COALESCE(sha.AtiendeDomingo, 1) END
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
            CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraApertura, CAST('08:00' AS TIME)) ELSE COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) END,
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80)),
            CAST(0 AS DECIMAL(10,2))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND
          (
              CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraApertura, CAST('08:00' AS TIME)) ELSE COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) END > CAST('00:00' AS TIME)
              OR CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraCierre, CAST('23:00' AS TIME)) ELSE COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) END < CAST('23:59' AS TIME)
          )

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
            CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraCierre, CAST('23:00' AS TIME)) ELSE COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) END,
            CAST('23:59' AS TIME),
            NULL,
            CAST(N'#64748b' AS NVARCHAR(20)),
            e.Id,
            e.Nombre,
            s.Nombre,
            CAST(N'Sede sin atencion (fuera de horario)' AS NVARCHAR(200)),
            CAST(N'BLOQUEADO_NO_ATENCION' AS NVARCHAR(40)),
            CAST(N'Bloqueado/No atencion' AS NVARCHAR(80)),
            CAST(0 AS DECIMAL(10,2))
        FROM Fechas f
        INNER JOIN dbo.Sedes s ON s.NegocioId = @NegocioId
        INNER JOIN dbo.EspaciosDeportivos e ON e.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
        LEFT JOIN dbo.SedeFechasInhabilitadas sfi ON sfi.SedeId = s.Id AND sfi.Activo = 1 AND sfi.Fecha = f.Fecha
        WHERE (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND sfi.SedeId IS NULL
          AND
          (
              CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraApertura, CAST('08:00' AS TIME)) ELSE COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) END > CAST('00:00' AS TIME)
              OR CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraCierre, CAST('23:00' AS TIME)) ELSE COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) END < CAST('23:59' AS TIME)
          )
        ORDER BY Fecha, HoraInicio
        OPTION (MAXRECURSION 1000);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
