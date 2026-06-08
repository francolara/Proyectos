
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_Actualizar]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 79)
-- Firma: Codex - 07/04/2026 | Edita reserva con comentario, valida politica de pago por estado, auto-ajusta a Pagada cuando el pago acumulado llega al 100%, y registra hasta 2 pagos por reserva con fecha de pago y numero de operacion alfanumerico opcional.
-- Firma: Codex - 10/04/2026 | Bloquea cambio a Cancelada cuando la reserva tiene pagos registrados.
-- Firma: FRANCO LARA - 26/05/2026 | Prioriza horario configurable por espacio deportivo; si no aplica, usa horario de la sede.
-- Firma: FRANCO LARA - 06/06/2026 | Valida cruces usando el espacio reservado y sus espacios compartidos activos.
-- Firma: FRANCO LARA - 08/06/2026 | Distingue bloqueo directo y espacios compuestos para evitar sobrebloqueos por propagacion en cadena.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_Actualizar]
    @Id INT,
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ClienteId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @Estado INT,
    @RegistrarPago BIT = 0,
    @FormaPagoId INT = NULL,
    @FechaPago DATETIME2 = NULL,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Comentario NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF @Total <= 0
            RAISERROR('El precio de espacio debe ser mayor que cero.', 16, 1);

        IF @Adelanto < 0
            RAISERROR('El monto de pago no puede ser negativo.', 16, 1);

        SET @NumeroOperacion = NULLIF(LTRIM(RTRIM(@NumeroOperacion)), N'');
        IF @NumeroOperacion IS NOT NULL AND @NumeroOperacion LIKE N'%[^0-9A-Za-z]%'
            RAISERROR('El numero de operacion solo puede contener caracteres alfanumericos.', 16, 1);

        DECLARE @AdelantoActual DECIMAL(10,2);
        DECLARE @AdelantoFinal DECIMAL(10,2);
        DECLARE @PagoNuevo DECIMAL(10,2);
        DECLARE @ConteoPagos INT;
        DECLARE @PoliticaConfirmacionPago TINYINT;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2);
        DECLARE @PagoMinimoRequerido DECIMAL(10,2);
        DECLARE @EstadoFinal INT;

        SELECT @AdelantoActual = r.Adelanto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @AdelantoActual IS NULL
            RAISERROR('No se encontro la reserva para actualizar.', 16, 1);

        SET @PagoNuevo = CASE WHEN @RegistrarPago = 1 THEN @Adelanto ELSE 0 END;
        SET @AdelantoFinal = @AdelantoActual + @PagoNuevo;

        IF @RegistrarPago = 1
        BEGIN
            IF @PagoNuevo <= 0
                RAISERROR('Debes ingresar un pago mayor que cero para registrar adelanto/pago.', 16, 1);

            IF @FormaPagoId IS NULL OR @FormaPagoId <= 0
                RAISERROR('Selecciona una forma de pago para registrar el adelanto/pago.', 16, 1);

            IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago fp WHERE fp.Id = @FormaPagoId AND fp.NegocioId = @NegocioId AND fp.Activo = 1)
                RAISERROR('La forma de pago seleccionada no es valida para el negocio.', 16, 1);

            IF @FechaPago IS NULL
                SET @FechaPago = CAST(SYSUTCDATETIME() AS DATE);

            IF CAST(@FechaPago AS DATE) > CAST(SYSUTCDATETIME() AS DATE)
                RAISERROR('La fecha de pago no puede ser mayor al dia actual.', 16, 1);

            SELECT @ConteoPagos = COUNT(1) FROM dbo.Pagos WHERE ReservaId = @Id;
            IF COALESCE(@ConteoPagos, 0) >= 2
                RAISERROR('La reserva ya tiene 2 pagos registrados. No se pueden registrar mas pagos.', 16, 1);
        END

        IF @AdelantoFinal > @Total
            RAISERROR('La suma de pagos excede el total de la reserva.', 16, 1);

        SELECT
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId;

        IF @PoliticaConfirmacionPago NOT IN (0, 1, 2)
            SET @PoliticaConfirmacionPago = 0;

        SET @EstadoFinal = @Estado;
        IF @AdelantoFinal >= @Total AND @Estado NOT IN (5, 6)
            SET @EstadoFinal = 4;

        IF @EstadoFinal = 5
        BEGIN
            IF @RegistrarPago = 1 OR EXISTS (SELECT 1 FROM dbo.Pagos p WHERE p.ReservaId = @Id AND COALESCE(p.Monto, 0) > 0)
                RAISERROR('No se puede cancelar la reserva porque tiene pagos registrados. Elimina los pagos para continuar.', 16, 1);
        END

        IF @EstadoFinal = 4
        BEGIN
            IF @AdelantoFinal < @Total
                RAISERROR('Para marcar como pagada, el pago acumulado debe ser 100% del precio del espacio.', 16, 1);
        END

        IF @EstadoFinal = 2
        BEGIN
            IF @PoliticaConfirmacionPago = 1
            BEGIN
                IF @PorcentajeAdelantoMinimo IS NULL OR @PorcentajeAdelantoMinimo <= 0 OR @PorcentajeAdelantoMinimo > 100
                    RAISERROR('La configuracion del porcentaje minimo de adelanto no es valida para confirmar.', 16, 1);

                SET @PagoMinimoRequerido = ROUND(@Total * (@PorcentajeAdelantoMinimo / 100.0), 2);
                IF @AdelantoFinal < @PagoMinimoRequerido
                    RAISERROR('No se puede confirmar: falta alcanzar el adelanto minimo configurado.', 16, 1);
            END
            ELSE IF @PoliticaConfirmacionPago = 2
            BEGIN
                IF @AdelantoFinal < @Total
                    RAISERROR('No se puede confirmar: la politica del negocio exige pago total (100%).', 16, 1);
            END
        END

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

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId IN (SELECT EspacioDeportivoId FROM @EspaciosAfectados)
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND r.Id <> @Id
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Cruce de horario detectado.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.BloqueosHorario b
            WHERE b.EspacioDeportivoId IN (SELECT EspacioDeportivoId FROM @EspaciosAfectados)
              AND b.Fecha = @Fecha
              AND b.Activo = 1
              AND @HoraInicio < b.HoraFin
              AND @HoraFin > b.HoraInicio
        )
            RAISERROR('El horario esta bloqueado para ese espacio.', 16, 1);

        BEGIN TRANSACTION;

        IF @RegistrarPago = 1 AND @PagoNuevo > 0
        BEGIN
            INSERT INTO dbo.Pagos
            (
                ReservaId, FechaPago, Monto, FormaPago, NumeroOperacion, Observacion,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @Id, @FechaPago, @PagoNuevo, @FormaPagoId, @NumeroOperacion, N'Pago registrado en edicion de reserva.',
                SYSUTCDATETIME(), @Usuario
            );
        END

        UPDATE r
        SET r.EspacioDeportivoId = @EspacioDeportivoId,
            r.ClienteId = @ClienteId,
            r.Fecha = @Fecha,
            r.HoraInicio = @HoraInicio,
            r.HoraFin = @HoraFin,
            r.Total = @Total,
            r.Adelanto = @AdelantoFinal,
            r.Saldo = (@Total - @AdelantoFinal),
            r.Estado = @EstadoFinal,
            r.Comentario = NULLIF(LTRIM(RTRIM(@Comentario)), N''),
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la reserva para actualizar.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'EDIT', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO


