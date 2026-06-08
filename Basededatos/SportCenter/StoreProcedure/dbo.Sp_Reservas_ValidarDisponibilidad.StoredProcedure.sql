
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_ValidarDisponibilidad]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 29_Reservas_ValidarDisponibilidad_Modal.sql (linea 8)
-- Firma: FRANCO LARA - 26/05/2026 | Prioriza horario configurable por espacio deportivo; si no aplica, usa horario de la sede.
-- Firma: FRANCO LARA - 06/06/2026 | Valida cruces usando el espacio reservado y sus espacios compartidos activos.
-- Firma: FRANCO LARA - 08/06/2026 | Distingue bloqueo directo y espacios compuestos para evitar sobrebloqueos por propagacion en cadena.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_ValidarDisponibilidad]
    @NegocioId INT,
    @ReservaId INT = NULL,
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'La hora fin debe ser mayor que la hora inicio.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        DECLARE @SedeId INT, @HoraApertura TIME, @HoraCierre TIME;
        DECLARE @AtiendeLunes BIT, @AtiendeMartes BIT, @AtiendeMiercoles BIT, @AtiendeJueves BIT, @AtiendeViernes BIT, @AtiendeSabado BIT, @AtiendeDomingo BIT;
        DECLARE @DiaSemana INT, @DiaHabilitado BIT;
        DECLARE @EspaciosAfectados TABLE (EspacioDeportivoId INT NOT NULL PRIMARY KEY);

        SELECT
            @SedeId = s.Id,
            @HoraApertura = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraApertura, CAST('08:00' AS TIME)) ELSE COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) END,
            @HoraCierre = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.HoraCierre, CAST('23:00' AS TIME)) ELSE COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) END,
            @AtiendeLunes = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeLunes, 1) ELSE COALESCE(sha.AtiendeLunes, 1) END,
            @AtiendeMartes = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeMartes, 1) ELSE COALESCE(sha.AtiendeMartes, 1) END,
            @AtiendeMiercoles = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeMiercoles, 1) ELSE COALESCE(sha.AtiendeMiercoles, 1) END,
            @AtiendeJueves = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeJueves, 1) ELSE COALESCE(sha.AtiendeJueves, 1) END,
            @AtiendeViernes = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeViernes, 1) ELSE COALESCE(sha.AtiendeViernes, 1) END,
            @AtiendeSabado = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeSabado, 1) ELSE COALESCE(sha.AtiendeSabado, 1) END,
            @AtiendeDomingo = CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeDomingo, 1) ELSE COALESCE(sha.AtiendeDomingo, 1) END
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
        WHERE e.Id = @EspacioDeportivoId
          AND s.NegocioId = @NegocioId
          AND e.Estado = 1;

        IF @SedeId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'El espacio deportivo no esta disponible para este negocio.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        IF EXISTS (SELECT 1 FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = @SedeId AND sfi.Fecha = @Fecha AND sfi.Activo = 1)
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'La sede no atiende en la fecha seleccionada.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        SET @DiaSemana = (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1;
        SET @DiaHabilitado = CASE @DiaSemana
            WHEN 1 THEN @AtiendeLunes
            WHEN 2 THEN @AtiendeMartes
            WHEN 3 THEN @AtiendeMiercoles
            WHEN 4 THEN @AtiendeJueves
            WHEN 5 THEN @AtiendeViernes
            WHEN 6 THEN @AtiendeSabado
            WHEN 7 THEN @AtiendeDomingo
            ELSE 0 END;

        IF COALESCE(@DiaHabilitado, 0) = 0
        BEGIN
            SELECT CAST(0 AS BIT) AS Disponible, CAST(N'La sede no atiende el dia seleccionado.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        IF @HoraInicio < @HoraApertura OR @HoraFin > @HoraCierre
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        N'Horario fuera de atencion. La sede atiende de ',
                        CONVERT(NVARCHAR(5), @HoraApertura, 108),
                        N' a ',
                        CONVERT(NVARCHAR(5), @HoraCierre, 108),
                        N'.'
                    )
                    AS NVARCHAR(300)
                ) AS Mensaje,
                CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo,
                CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        INSERT INTO @EspaciosAfectados (EspacioDeportivoId)
        VALUES (@EspacioDeportivoId);

        INSERT INTO @EspaciosAfectados (EspacioDeportivoId)
        SELECT DISTINCT ec.EspacioRelacionadoId
        FROM dbo.EspaciosDeportivosCompartidos ec
        INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioRelacionadoId
        WHERE ec.EspacioDeportivoId = @EspacioDeportivoId
          AND ec.Activo = 1
          AND ec.TipoRelacion = N'DIRECTO'
          AND er.Estado = 1
          AND NOT EXISTS (SELECT 1 FROM @EspaciosAfectados ea WHERE ea.EspacioDeportivoId = ec.EspacioRelacionadoId);

        INSERT INTO @EspaciosAfectados (EspacioDeportivoId)
        SELECT DISTINCT ec.EspacioRelacionadoId
        FROM dbo.EspaciosDeportivosCompartidos ec
        INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioRelacionadoId
        WHERE ec.EspacioDeportivoId = @EspacioDeportivoId
          AND ec.Activo = 1
          AND ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
          AND er.Estado = 1
          AND NOT EXISTS (SELECT 1 FROM @EspaciosAfectados ea WHERE ea.EspacioDeportivoId = ec.EspacioRelacionadoId);

        INSERT INTO @EspaciosAfectados (EspacioDeportivoId)
        SELECT DISTINCT ec.EspacioDeportivoId
        FROM dbo.EspaciosDeportivosCompartidos ec
        INNER JOIN dbo.EspaciosDeportivos er ON er.Id = ec.EspacioDeportivoId
        WHERE ec.EspacioRelacionadoId = @EspacioDeportivoId
          AND ec.Activo = 1
          AND ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
          AND er.Estado = 1
          AND NOT EXISTS (SELECT 1 FROM @EspaciosAfectados ea WHERE ea.EspacioDeportivoId = ec.EspacioDeportivoId);

        INSERT INTO @EspaciosAfectados (EspacioDeportivoId)
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
          AND NOT EXISTS (SELECT 1 FROM @EspaciosAfectados ea WHERE ea.EspacioDeportivoId = ed.EspacioRelacionadoId);

        INSERT INTO @EspaciosAfectados (EspacioDeportivoId)
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
          AND NOT EXISTS (SELECT 1 FROM @EspaciosAfectados ea WHERE ea.EspacioDeportivoId = ep.EspacioDeportivoId);

        DECLARE @ReservaCruceId INT = NULL, @ReservaCruceInicio TIME = NULL, @ReservaCruceFin TIME = NULL, @ReservaCruceEspacio NVARCHAR(150) = NULL;
        SELECT TOP 1
            @ReservaCruceId = r.Id,
            @ReservaCruceInicio = r.HoraInicio,
            @ReservaCruceFin = r.HoraFin,
            @ReservaCruceEspacio = e.Nombre
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        WHERE r.EspacioDeportivoId IN (SELECT EspacioDeportivoId FROM @EspaciosAfectados)
          AND r.Fecha = @Fecha
          AND r.Estado NOT IN (5, 6)
          AND (@ReservaId IS NULL OR r.Id <> @ReservaId)
          AND @HoraInicio < r.HoraFin
          AND @HoraFin > r.HoraInicio
        ORDER BY CASE WHEN r.EspacioDeportivoId = @EspacioDeportivoId THEN 0 ELSE 1 END, r.HoraInicio;

        IF @ReservaCruceId IS NOT NULL
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        CASE WHEN @ReservaCruceEspacio IS NOT NULL AND EXISTS (SELECT 1 FROM @EspaciosAfectados WHERE EspacioDeportivoId <> @EspacioDeportivoId) AND @ReservaCruceEspacio <> N'' THEN N'Cruce con reserva en ' + @ReservaCruceEspacio + N' #' ELSE N'Cruce con reserva #' END,
                        @ReservaCruceId,
                        N' (',
                        CONVERT(NVARCHAR(5), @ReservaCruceInicio, 108),
                        N' - ',
                        CONVERT(NVARCHAR(5), @ReservaCruceFin, 108),
                        N').'
                    )
                    AS NVARCHAR(300)
                ) AS Mensaje,
                CAST(N'RESERVA' AS NVARCHAR(20)) AS ConflictoTipo,
                @ReservaCruceId AS ConflictoId;
            RETURN;
        END;

        DECLARE @BloqueoInicio TIME = NULL, @BloqueoFin TIME = NULL, @BloqueoMotivo NVARCHAR(250) = NULL, @BloqueoEspacio NVARCHAR(150) = NULL;
        SELECT TOP 1
            @BloqueoInicio = b.HoraInicio,
            @BloqueoFin = b.HoraFin,
            @BloqueoMotivo = b.Motivo,
            @BloqueoEspacio = e.Nombre
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        WHERE b.EspacioDeportivoId IN (SELECT EspacioDeportivoId FROM @EspaciosAfectados)
          AND b.Fecha = @Fecha
          AND b.Activo = 1
          AND @HoraInicio < b.HoraFin
          AND @HoraFin > b.HoraInicio
        ORDER BY CASE WHEN b.EspacioDeportivoId = @EspacioDeportivoId THEN 0 ELSE 1 END, b.HoraInicio;

        IF @BloqueoInicio IS NOT NULL
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        N'Horario bloqueado (',
                        CONVERT(NVARCHAR(5), @BloqueoInicio, 108),
                        N' - ',
                        CONVERT(NVARCHAR(5), @BloqueoFin, 108),
                        N') en ',
                        COALESCE(@BloqueoEspacio, N'espacio relacionado'),
                        N'. Motivo: ',
                        COALESCE(@BloqueoMotivo, N'Sin detalle')
                    )
                    AS NVARCHAR(300)
                ) AS Mensaje,
                CAST(N'BLOQUEO' AS NVARCHAR(20)) AS ConflictoTipo,
                CAST(NULL AS INT) AS ConflictoId;
            RETURN;
        END;

        SELECT CAST(1 AS BIT) AS Disponible, CAST(N'Horario disponible.' AS NVARCHAR(300)) AS Mensaje, CAST(NULL AS NVARCHAR(20)) AS ConflictoTipo, CAST(NULL AS INT) AS ConflictoId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
