USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_Crear]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 26_Reservas_Validacion_Horario_Sede.sql (linea 9)
-- Firma: Codex - 07/04/2026 | Crea reserva con politica de pago del negocio, registro opcional de pago en creacion y comentario, incluyendo fecha de pago y numero de operacion alfanumerico opcional.
-- Firma: Codex - 14/04/2026 | Agrega CanalOrigen en reservas, genera notificacion para origen CLIENTE_WEB y ajusta concatenacion compatible en mensaje/url.
-- Firma: Codex - 17/04/2026 | Aisla notificacion en flujo CLIENTE_WEB sin INSERT-EXEC, agrega salida opcional @ReservaId y mantiene SELECT final del Id real creado.
-- Firma: FRANCO LARA - 03/05/2026 | Permite aplicar cupon por codigo en reserva (admin/web), valida vigencia/limite por sede-espacio y registra uso acumulado. Ajusta validacion de adelanto/estado para evaluar el total final luego del cupon.
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
    @FechaPago DATETIME2 = NULL,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Comentario NVARCHAR(500) = NULL,
    @CodigoCupon NVARCHAR(30) = NULL,
    @CanalOrigen NVARCHAR(20) = N'ADMIN',
    @ReservaId INT = NULL OUTPUT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF @Total <= 0
            RAISERROR('El precio de espacio debe ser mayor que cero.', 16, 1);

        DECLARE @TotalOriginal DECIMAL(10,2) = @Total;
        DECLARE @DescuentoCupon DECIMAL(10,2) = 0;
        DECLARE @CuponId INT = NULL;

        IF @Adelanto < 0
            RAISERROR('El monto de pago no puede ser negativo.', 16, 1);

        IF @RegistrarPago = 0
            SET @Adelanto = 0;

        IF @RegistrarPago = 1 AND @Adelanto <= 0
            RAISERROR('Debes ingresar un pago mayor que cero para registrar adelanto/pago.', 16, 1);

        SET @NumeroOperacion = NULLIF(LTRIM(RTRIM(@NumeroOperacion)), N'');
        IF @NumeroOperacion IS NOT NULL AND @NumeroOperacion LIKE N'%[^0-9A-Za-z]%'
            RAISERROR('El numero de operacion solo puede contener caracteres alfanumericos.', 16, 1);

        IF @RegistrarPago = 1
        BEGIN
            IF @FormaPagoId IS NULL OR @FormaPagoId <= 0
                RAISERROR('Selecciona una forma de pago para registrar el adelanto/pago.', 16, 1);

            IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago fp WHERE fp.Id = @FormaPagoId AND fp.NegocioId = @NegocioId AND fp.Activo = 1)
                RAISERROR('La forma de pago seleccionada no es valida para el negocio.', 16, 1);

            IF @FechaPago IS NULL
                SET @FechaPago = CAST(SYSUTCDATETIME() AS DATE);

            IF CAST(@FechaPago AS DATE) > CAST(SYSUTCDATETIME() AS DATE)
                RAISERROR('La fecha de pago no puede ser mayor al dia actual.', 16, 1);
        END

        DECLARE @PoliticaConfirmacionPago TINYINT;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2);
        DECLARE @PagoMinimoRequerido DECIMAL(10,2);
        DECLARE @EstadoCalculado INT;

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

        SET @CodigoCupon = UPPER(NULLIF(LTRIM(RTRIM(@CodigoCupon)), N''));
        IF @CodigoCupon IS NOT NULL
        BEGIN
            DECLARE @HoyCupon DATE = CAST(SYSUTCDATETIME() AS DATE);
            SELECT TOP 1
                @CuponId = c.Id,
                @DescuentoCupon = CASE
                    WHEN c.TipoDescuento = N'PORCENTAJE' THEN ROUND(@TotalOriginal * (c.ValorDescuento / 100.0), 2)
                    ELSE c.ValorDescuento
                END
            FROM dbo.Cupones c WITH (UPDLOCK, HOLDLOCK)
            WHERE c.NegocioId = @NegocioId
              AND c.CodigoCupon = @CodigoCupon
              AND c.Activo = 1
              AND c.FechaInicio <= @HoyCupon
              AND c.FechaFin >= @HoyCupon
              AND c.CantidadUsosActuales < c.CantidadMaxUsos
              AND (c.SedeId IS NULL OR c.SedeId = @SedeId)
              AND (c.EspacioDeportivoId IS NULL OR c.EspacioDeportivoId = @EspacioDeportivoId);

            IF @CuponId IS NULL
                RAISERROR('El cupon no es valido para la reserva seleccionada.', 16, 1);

            IF @DescuentoCupon < 0 SET @DescuentoCupon = 0;
            IF @DescuentoCupon > @TotalOriginal SET @DescuentoCupon = @TotalOriginal;
            SET @Total = @TotalOriginal - @DescuentoCupon;
        END

        IF @Adelanto > @Total
            RAISERROR('El adelanto/pago no puede ser mayor que el total.', 16, 1);

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

        BEGIN TRANSACTION;

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin, Estado,
            Total, Adelanto, Saldo, Comentario, CodigoCuponAplicado, DescuentoCupon, CanalOrigen, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin, @EstadoCalculado,
            @Total, @Adelanto, (@Total - @Adelanto), NULLIF(LTRIM(RTRIM(@Comentario)), N''), @CodigoCupon, @DescuentoCupon, COALESCE(NULLIF(LTRIM(RTRIM(@CanalOrigen)), N''), N'ADMIN'), SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT = SCOPE_IDENTITY();
        SET @ReservaId = @Id;

        IF @RegistrarPago = 1 AND @Adelanto > 0
        BEGIN
            INSERT INTO dbo.Pagos
            (
                ReservaId, FechaPago, Monto, FormaPago, NumeroOperacion, Observacion,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @Id, @FechaPago, @Adelanto, @FormaPagoId, @NumeroOperacion, N'Pago registrado al crear reserva.',
                SYSUTCDATETIME(), @Usuario
            );
        END

        IF @CuponId IS NOT NULL
        BEGIN
            UPDATE dbo.Cupones
            SET CantidadUsosActuales = CantidadUsosActuales + 1,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
            WHERE Id = @CuponId;

            INSERT INTO dbo.CuponesUso
            (
                CuponId, ReservaId, ClienteId, MontoAntes, MontoDescuento, MontoFinal, CanalOrigen, FechaUso, UsuarioCreacion
            )
            VALUES
            (
                @CuponId, @Id, @ClienteId, @TotalOriginal, @DescuentoCupon, @Total, COALESCE(NULLIF(LTRIM(RTRIM(@CanalOrigen)), N''), N'ADMIN'), SYSUTCDATETIME(), @Usuario
            );
        END

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        IF UPPER(COALESCE(NULLIF(LTRIM(RTRIM(@CanalOrigen)), N''), N'ADMIN')) = N'CLIENTE_WEB'
        BEGIN
            DECLARE @MensajeNotificacion NVARCHAR(300);
            DECLARE @UrlNotificacion NVARCHAR(300);
            SET @MensajeNotificacion = N'Reserva #' + CONVERT(NVARCHAR(20), @Id) + N' creada desde portal cliente.';
            SET @UrlNotificacion = N'/Reservas?negocioId=' + CONVERT(NVARCHAR(20), @NegocioId);
            EXEC dbo.Sp_Notificaciones_Crear
                @NegocioId = @NegocioId,
                @Tipo = N'RESERVA_CLIENTE_WEB',
                @Titulo = N'Nueva reserva web',
                @Mensaje = @MensajeNotificacion,
                @Entidad = N'Reserva',
                @EntidadId = @Id,
                @UrlDestino = @UrlNotificacion,
                @DevolverResultado = 0;
        END


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
