/*
Firma: Codex - 07/04/2026
Descripcion: Reserva pop-up con politica de pago, cotizacion, comentario y limite maximo de 2 pagos por reserva.
Firma: FRANCO LARA - 20/05/2026
Descripcion: El tramo 23:00-23:59 se factura como hora completa (60 min) en cotizacion.
*/
USE [DbSportCenter]
GO
IF COL_LENGTH('dbo.Reservas', 'Comentario') IS NULL
BEGIN
    ALTER TABLE dbo.Reservas ADD Comentario NVARCHAR(500) NULL;
END
GO

USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 07/04/2026 | Cotiza precio de reserva segun tarifa horaria, promociones activas y politica de pago del negocio.
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Cotizar
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        DECLARE @DuracionMinutos INT = DATEDIFF(MINUTE, @HoraInicio, @HoraFin);
        IF @DuracionMinutos NOT IN (30, 60)
           AND NOT (@HoraInicio = '23:00:00' AND @HoraFin = '23:59:00')
            RAISERROR('Solo se permite reservas de 30 o 60 minutos.', 16, 1);
        DECLARE @DuracionFacturableMinutos INT =
            CASE
                WHEN @HoraInicio = '23:00:00' AND @HoraFin = '23:59:00' THEN 60
                ELSE @DuracionMinutos
            END;

        DECLARE @SedeId INT;
        DECLARE @DiaSemana INT = (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1;

        SELECT @SedeId = s.Id
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @EspacioDeportivoId
          AND s.NegocioId = @NegocioId
          AND e.Estado = 1
          AND s.Activo = 1;

        IF @SedeId IS NULL
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        DECLARE @PrecioHora DECIMAL(10,2);
        SELECT TOP 1
            @PrecioHora = t.Precio
        FROM dbo.Tarifas t
        WHERE t.EspacioDeportivoId = @EspacioDeportivoId
          AND t.Activa = 1
          AND t.DiaSemana = @DiaSemana
          AND @HoraInicio >= t.HoraInicio
          AND @HoraFin <= t.HoraFin
        ORDER BY t.HoraInicio DESC;

        IF @PrecioHora IS NULL
            RAISERROR('No existe tarifa configurada para el horario seleccionado.', 16, 1);

        DECLARE @PrecioBase DECIMAL(10,2) = ROUND(@PrecioHora * (@DuracionFacturableMinutos / 60.0), 2);
        DECLARE @DescuentoPct DECIMAL(5,2) = 0;

        SELECT TOP 1
            @DescuentoPct = p.PorcentajeDescuento
        FROM dbo.PromocionesHorario p
        WHERE p.NegocioId = @NegocioId
          AND p.Activo = 1
          AND @Fecha BETWEEN p.FechaInicio AND p.FechaFin
          AND @HoraInicio >= p.HoraInicio
          AND @HoraFin <= p.HoraFin
          AND (p.EspacioDeportivoId IS NULL OR p.EspacioDeportivoId = @EspacioDeportivoId)
          AND (p.SedeId IS NULL OR p.SedeId = @SedeId)
        ORDER BY
            CASE
                WHEN p.EspacioDeportivoId = @EspacioDeportivoId THEN 3
                WHEN p.SedeId = @SedeId THEN 2
                ELSE 1
            END DESC,
            p.PorcentajeDescuento DESC,
            p.Id DESC;

        DECLARE @PrecioFinal DECIMAL(10,2) = ROUND(@PrecioBase * (1 - (COALESCE(@DescuentoPct, 0) / 100.0)), 2);

        DECLARE @MonedaNombre NVARCHAR(80) = N'PEN';
        DECLARE @MonedaSimbolo NVARCHAR(10) = N'S/';
        DECLARE @PoliticaConfirmacionPago TINYINT = 0;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL;

        SELECT
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo,
            @MonedaNombre = COALESCE(m.Nombre, N'PEN'),
            @MonedaSimbolo = COALESCE(NULLIF(LTRIM(RTRIM(m.Simbolo)), N''), N'S/')
        FROM dbo.Negocios n
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        WHERE n.Id = @NegocioId;

        SELECT
            CAST(N'Tarifa calculada correctamente.' AS NVARCHAR(200)) AS Mensaje,
            @PrecioBase AS PrecioBase,
            COALESCE(@DescuentoPct, 0) AS DescuentoPct,
            @PrecioFinal AS PrecioFinal,
            @MonedaSimbolo AS MonedaSimbolo,
            @MonedaNombre AS MonedaNombre,
            CAST(COALESCE(@PoliticaConfirmacionPago, 0) AS TINYINT) AS PoliticaConfirmacionPago,
            @PorcentajeAdelantoMinimo AS PorcentajeAdelantoMinimo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_Crear]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 26_Reservas_Validacion_Horario_Sede.sql (linea 9)
-- Firma: Codex - 07/04/2026 | Crea reserva con politica de pago del negocio, registro opcional de pago en creacion y comentario.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_Crear]
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

        IF @RegistrarPago = 0
            SET @Adelanto = 0;

        IF @RegistrarPago = 1 AND @Adelanto <= 0
            RAISERROR('Debes ingresar un pago mayor que cero para registrar adelanto/pago.', 16, 1);

        IF @Adelanto > @Total
            RAISERROR('El adelanto/pago no puede ser mayor que el total.', 16, 1);

        IF @RegistrarPago = 1
        BEGIN
            IF @FormaPagoId IS NULL OR @FormaPagoId <= 0
                RAISERROR('Selecciona una forma de pago para registrar el adelanto/pago.', 16, 1);

            IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago fp WHERE fp.Id = @FormaPagoId AND fp.NegocioId = @NegocioId AND fp.Activo = 1)
                RAISERROR('La forma de pago seleccionada no es valida para el negocio.', 16, 1);
        END

        DECLARE @PoliticaConfirmacionPago TINYINT;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2);
        DECLARE @PagoMinimoRequerido DECIMAL(10,2);
        DECLARE @EstadoCalculado INT;

        SELECT
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId;

        IF @PoliticaConfirmacionPago NOT IN (0, 1, 2)
            SET @PoliticaConfirmacionPago = 0;

        SET @EstadoCalculado = 1;
        IF @Adelanto >= @Total
            SET @EstadoCalculado = 4;
        ELSE IF @Adelanto > 0
        BEGIN
            IF @PoliticaConfirmacionPago = 0
                SET @EstadoCalculado = 2;
            ELSE IF @PoliticaConfirmacionPago = 1
            BEGIN
                IF @PorcentajeAdelantoMinimo IS NULL OR @PorcentajeAdelantoMinimo <= 0 OR @PorcentajeAdelantoMinimo > 100
                    RAISERROR('La configuracion del porcentaje minimo de adelanto no es valida.', 16, 1);

                SET @PagoMinimoRequerido = ROUND(@Total * (@PorcentajeAdelantoMinimo / 100.0), 2);
                IF @Adelanto >= @PagoMinimoRequerido
                    SET @EstadoCalculado = 2;
            END
        END

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

        BEGIN TRANSACTION;

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin, Estado,
            Total, Adelanto, Saldo, Comentario, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin, @EstadoCalculado,
            @Total, @Adelanto, (@Total - @Adelanto), NULLIF(LTRIM(RTRIM(@Comentario)), N''), SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT = SCOPE_IDENTITY();

        IF @RegistrarPago = 1 AND @Adelanto > 0
        BEGIN
            INSERT INTO dbo.Pagos
            (
                ReservaId, FechaPago, Monto, FormaPago, NumeroOperacion, Observacion,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @Id, SYSUTCDATETIME(), @Adelanto, @FormaPagoId, NULL, N'Pago registrado al crear reserva.',
                SYSUTCDATETIME(), @Usuario
            );
        END

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
        SELECT @Id;
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

USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_Actualizar]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 79)
-- Firma: Codex - 07/04/2026 | Edita reserva con comentario, valida politica de pago por estado y registra hasta 2 pagos por reserva.
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

        DECLARE @AdelantoActual DECIMAL(10,2);
        DECLARE @AdelantoFinal DECIMAL(10,2);
        DECLARE @PagoNuevo DECIMAL(10,2);
        DECLARE @ConteoPagos INT;
        DECLARE @PoliticaConfirmacionPago TINYINT;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2);
        DECLARE @PagoMinimoRequerido DECIMAL(10,2);

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

        IF @Estado = 4
        BEGIN
            IF @AdelantoFinal < @Total
                RAISERROR('Para marcar como pagada, el pago acumulado debe ser 100% del precio del espacio.', 16, 1);
        END

        IF @Estado = 2
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

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
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
            WHERE b.EspacioDeportivoId = @EspacioDeportivoId
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
                @Id, SYSUTCDATETIME(), @PagoNuevo, @FormaPagoId, NULL, N'Pago registrado en edicion de reserva.',
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
            r.Estado = @Estado,
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

USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_ObtenerPorId]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 104)
-- Firma: Codex - 07/04/2026 | Incluye Comentario en consulta de detalle de reserva para pop-up.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id,
            r.EspacioDeportivoId,
            r.ClienteId,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            r.Adelanto,
            r.Estado,
            r.Comentario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Pagos_Crear]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 35_Maestros_FormasPago.sql (linea 186)
-- Firma: Codex - 07/04/2026 | Limita a maximo 2 pagos por reserva y ajusta estado automatico (confirmada/pagada) segun pago acumulado.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Crear]
    @NegocioId INT,
    @ReservaId INT,
    @FechaPago DATETIME2,
    @Monto DECIMAL(10,2),
    @FormaPago INT,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Observacion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Monto <= 0
            RAISERROR('El monto debe ser mayor que cero.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE Id = @FormaPago AND Activo = 1)
            RAISERROR('La forma de pago no es valida.', 16, 1);

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoActual DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);
        DECLARE @CantidadPagos INT;

        SELECT @TotalReserva = r.Total
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @TotalReserva IS NULL
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        SELECT
            @PagadoActual = COALESCE(SUM(p.Monto), 0),
            @CantidadPagos = COUNT(1)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        IF COALESCE(@CantidadPagos, 0) >= 2
            RAISERROR('La reserva ya tiene 2 pagos registrados. No se pueden registrar mas pagos.', 16, 1);

        SET @NuevoPagado = @PagadoActual + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Pagos
        (
            ReservaId, FechaPago, Monto, FormaPago, NumeroOperacion, Observacion,
            FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @ReservaId, @FechaPago, @Monto, @FormaPago, @NumeroOperacion, @Observacion,
            SYSUTCDATETIME(), @Usuario
        );

        UPDATE r
        SET Adelanto = @NuevoPagado,
            Saldo = (r.Total - @NuevoPagado),
            Estado = CASE
                        WHEN (r.Total - @NuevoPagado) <= 0 THEN 4
                        WHEN @NuevoPagado > 0 AND r.Estado = 1 THEN 2
                        ELSE r.Estado
                     END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        WHERE r.Id = @ReservaId;

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'CREATE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
        SELECT @Id;
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
