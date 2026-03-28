-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Validacion previa de disponibilidad para modal de reservas (con detalle exacto de conflicto).
-- Firma:         Codex - 27/03/2026
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_ValidarDisponibilidad
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

        SELECT
            @SedeId = s.Id,
            @HoraApertura = COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)),
            @HoraCierre = COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)),
            @AtiendeLunes = COALESCE(sha.AtiendeLunes, 1),
            @AtiendeMartes = COALESCE(sha.AtiendeMartes, 1),
            @AtiendeMiercoles = COALESCE(sha.AtiendeMiercoles, 1),
            @AtiendeJueves = COALESCE(sha.AtiendeJueves, 1),
            @AtiendeViernes = COALESCE(sha.AtiendeViernes, 1),
            @AtiendeSabado = COALESCE(sha.AtiendeSabado, 1),
            @AtiendeDomingo = COALESCE(sha.AtiendeDomingo, 1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
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

        DECLARE @ReservaCruceId INT = NULL, @ReservaCruceInicio TIME = NULL, @ReservaCruceFin TIME = NULL;
        SELECT TOP 1
            @ReservaCruceId = r.Id,
            @ReservaCruceInicio = r.HoraInicio,
            @ReservaCruceFin = r.HoraFin
        FROM dbo.Reservas r
        WHERE r.EspacioDeportivoId = @EspacioDeportivoId
          AND r.Fecha = @Fecha
          AND r.Estado NOT IN (5, 6)
          AND (@ReservaId IS NULL OR r.Id <> @ReservaId)
          AND @HoraInicio < r.HoraFin
          AND @HoraFin > r.HoraInicio
        ORDER BY r.HoraInicio;

        IF @ReservaCruceId IS NOT NULL
        BEGIN
            SELECT
                CAST(0 AS BIT) AS Disponible,
                CAST(
                    CONCAT(
                        N'Cruce con reserva #',
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

        DECLARE @BloqueoInicio TIME = NULL, @BloqueoFin TIME = NULL, @BloqueoMotivo NVARCHAR(250) = NULL;
        SELECT TOP 1
            @BloqueoInicio = b.HoraInicio,
            @BloqueoFin = b.HoraFin,
            @BloqueoMotivo = b.Motivo
        FROM dbo.BloqueosHorario b
        WHERE b.EspacioDeportivoId = @EspacioDeportivoId
          AND b.Fecha = @Fecha
          AND b.Activo = 1
          AND @HoraInicio < b.HoraFin
          AND @HoraFin > b.HoraInicio
        ORDER BY b.HoraInicio;

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
                        N'). Motivo: ',
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
