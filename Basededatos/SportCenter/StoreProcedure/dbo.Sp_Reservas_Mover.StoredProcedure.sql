
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_Mover]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 26_Reservas_Validacion_Horario_Sede.sql (linea 93)
-- Firma: FRANCO LARA - 26/05/2026 | Prioriza horario configurable por espacio deportivo; si no aplica, usa horario de la sede.
-- Firma: FRANCO LARA - 06/06/2026 | Valida cruces usando el espacio reservado y sus espacios compartidos activos.
-- Firma: FRANCO LARA - 08/06/2026 | Distingue bloqueo directo y espacios compuestos para evitar sobrebloqueos por propagacion en cadena.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_Mover]
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
        DECLARE @EspaciosAfectados TABLE (EspacioDeportivoId INT NOT NULL PRIMARY KEY);

        SELECT
            @EspacioDeportivoId = r.EspacioDeportivoId,
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
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
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

        IF EXISTS (SELECT 1 FROM dbo.Reservas r WHERE r.EspacioDeportivoId IN (SELECT EspacioDeportivoId FROM @EspaciosAfectados) AND r.Fecha = @Fecha AND r.Estado NOT IN (5, 6) AND r.Id <> @Id AND @HoraInicio < r.HoraFin AND @HoraFin > r.HoraInicio)
            RAISERROR('Cruce de horario con otra reserva.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.BloqueosHorario b WHERE b.EspacioDeportivoId IN (SELECT EspacioDeportivoId FROM @EspaciosAfectados) AND b.Fecha = @Fecha AND b.Activo = 1 AND @HoraInicio < b.HoraFin AND @HoraFin > b.HoraInicio)
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

