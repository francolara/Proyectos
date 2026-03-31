-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Valida reservas segun dias/horario de sede y fechas no laborables.
-- Firma:         Codex - 27/03/2026
-- Firma:         Codex - 30/03/2026 | Sp_Reservas_Mover devuelve error cuando no encuentra reserva o no afecta filas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Crear
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ClienteId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

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
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = @SedeId AND sfi.Fecha = @Fecha AND sfi.Activo = 1)
            RAISERROR('La sede no atiende en la fecha seleccionada.', 16, 1);

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
            RAISERROR('La sede no atiende el dia seleccionado.', 16, 1);
        IF @HoraInicio < @HoraApertura OR @HoraFin > @HoraCierre
            RAISERROR('El horario de reserva esta fuera del horario de atencion de la sede.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Reservas r WHERE r.EspacioDeportivoId = @EspacioDeportivoId AND r.Fecha = @Fecha AND r.Estado NOT IN (5, 6) AND @HoraInicio < r.HoraFin AND @HoraFin > r.HoraInicio)
            RAISERROR('Cruce de horario detectado.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.BloqueosHorario b WHERE b.EspacioDeportivoId = @EspacioDeportivoId AND b.Fecha = @Fecha AND b.Activo = 1 AND @HoraInicio < b.HoraFin AND @HoraFin > b.HoraInicio)
            RAISERROR('El horario esta bloqueado para ese espacio.', 16, 1);

        INSERT INTO dbo.Reservas (EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin, Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion)
        VALUES (@EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin, @Estado, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Mover
    @NegocioId INT,
    @Id INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a la hora inicio.', 16, 1);

        DECLARE @EspacioDeportivoId INT, @SedeId INT, @HoraApertura TIME, @HoraCierre TIME;
        DECLARE @AtiendeLunes BIT, @AtiendeMartes BIT, @AtiendeMiercoles BIT, @AtiendeJueves BIT, @AtiendeViernes BIT, @AtiendeSabado BIT, @AtiendeDomingo BIT;
        DECLARE @DiaSemana INT, @DiaHabilitado BIT;

        SELECT
            @EspacioDeportivoId = r.EspacioDeportivoId,
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
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
          AND r.Estado NOT IN (5, 6);

        IF @EspacioDeportivoId IS NULL
            RAISERROR('No se encontro la reserva para mover.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = @SedeId AND sfi.Fecha = @Fecha AND sfi.Activo = 1)
            RAISERROR('La sede no atiende en la fecha seleccionada.', 16, 1);

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
            RAISERROR('La sede no atiende el dia seleccionado.', 16, 1);
        IF @HoraInicio < @HoraApertura OR @HoraFin > @HoraCierre
            RAISERROR('El horario de reserva esta fuera del horario de atencion de la sede.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Reservas r WHERE r.EspacioDeportivoId = @EspacioDeportivoId AND r.Fecha = @Fecha AND r.Estado NOT IN (5, 6) AND r.Id <> @Id AND @HoraInicio < r.HoraFin AND @HoraFin > r.HoraInicio)
            RAISERROR('Cruce de horario con otra reserva.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.BloqueosHorario b WHERE b.EspacioDeportivoId = @EspacioDeportivoId AND b.Fecha = @Fecha AND b.Activo = 1 AND @HoraInicio < b.HoraFin AND @HoraFin > b.HoraInicio)
            RAISERROR('El horario esta bloqueado para ese espacio.', 16, 1);

        UPDATE dbo.Reservas
        SET Fecha = @Fecha, HoraInicio = @HoraInicio, HoraFin = @HoraFin, FechaActualizacion = SYSUTCDATETIME(), UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la reserva para mover.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'MOVE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
