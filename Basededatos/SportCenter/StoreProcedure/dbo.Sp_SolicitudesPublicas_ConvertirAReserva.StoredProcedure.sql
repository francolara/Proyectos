
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_SolicitudesPublicas_ConvertirAReserva para tipo de documento SUNAT por defecto.
-- Firma: Codex - 06/04/2026 | Si la solicitud se convierte como Confirmada, valida politica de pago del negocio.
-- Firma: Codex - 06/04/2026 | Se elimina dependencia de NegocioClientes y se usa Clientes.NegocioId.
-- Firma: Codex - 14/04/2026 | Reserva generada desde solicitud publica persiste CanalOrigen=CLIENTE_WEB, crea notificacion para campanita admin y ajusta concatenacion compatible.
-- Firma: FRANCO LARA - 01/10/2026 | Asigna correlativos visibles por negocio al cliente y reserva creados desde una solicitud publica y devuelve ambos identificadores de la reserva.
-- Firma: FRANCO LARA - 06/10/2026 | Conserva CodigoMoneda canonico al convertir la solicitud en reserva.
CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_ConvertirAReserva
    @NegocioId INT,
    @Id INT,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @EstadoReserva INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Total < 0 OR @Adelanto < 0 OR @Adelanto > @Total
            RAISERROR('Montos invalidos para la conversion.', 16, 1);

        DECLARE @PoliticaConfirmacionPago TINYINT;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2);
        DECLARE @PagoMinimoRequerido DECIMAL(10,2);
        DECLARE @CodigoMoneda NVARCHAR(10);

        SELECT
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo,
            @CodigoMoneda = n.CodigoMoneda
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId;

        IF @CodigoMoneda IS NULL
            RAISERROR('El negocio debe configurar una moneda valida antes de convertir solicitudes.', 16, 1);

        IF @PoliticaConfirmacionPago NOT IN (0, 1, 2)
            SET @PoliticaConfirmacionPago = 0;

        IF @EstadoReserva = 2
        BEGIN
            IF @PoliticaConfirmacionPago = 1
            BEGIN
                IF @PorcentajeAdelantoMinimo IS NULL OR @PorcentajeAdelantoMinimo <= 0 OR @PorcentajeAdelantoMinimo > 100
                    RAISERROR('La configuracion del porcentaje minimo de adelanto no es valida para confirmar.', 16, 1);

                SET @PagoMinimoRequerido = ROUND(@Total * (@PorcentajeAdelantoMinimo / 100.0), 2);
                IF @Adelanto < @PagoMinimoRequerido
                    RAISERROR('No se puede confirmar: el adelanto no alcanza el porcentaje minimo configurado.', 16, 1);
            END
            ELSE IF @PoliticaConfirmacionPago = 2
            BEGIN
                IF @Adelanto < @Total
                    RAISERROR('No se puede confirmar: se requiere pago total (100%).', 16, 1);
            END
        END

        DECLARE @EspacioDeportivoId INT, @Fecha DATE, @HoraInicio TIME, @HoraFin TIME, @NombreSolicitante NVARCHAR(200), @Telefono NVARCHAR(30), @Correo NVARCHAR(200);
        DECLARE @ClienteId INT, @ReservaId INT;

        SELECT
            @EspacioDeportivoId = s.EspacioDeportivoId,
            @Fecha = s.Fecha,
            @HoraInicio = s.HoraInicio,
            @HoraFin = s.HoraFin,
            @NombreSolicitante = s.NombreSolicitante,
            @Telefono = s.Telefono,
            @Correo = s.Correo
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.Id = @Id
          AND se.NegocioId = @NegocioId
          AND s.Estado IN (1, 2);

        IF @EspacioDeportivoId IS NULL
            RAISERROR('Solicitud invalida para el negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('No se puede convertir: el horario ya fue tomado.', 16, 1);

        SELECT TOP (1) @ClienteId = c.Id
        FROM dbo.Clientes c
        WHERE c.NegocioId = @NegocioId
          AND c.Activo = 1
          AND c.NombresORazonSocial = @NombreSolicitante
          AND c.Telefono = @Telefono;

        BEGIN TRANSACTION;

        IF @ClienteId IS NULL
        BEGIN
            DECLARE @NumeroClientePorNegocio INT;
            EXEC dbo.Sp_NegocioCorrelativos_ObtenerSiguiente @NegocioId, N'CLIENTE', @Usuario, @NumeroClientePorNegocio OUTPUT;

            INSERT INTO dbo.Clientes
            (
                NumeroPorNegocio, NegocioId, NombresORazonSocial, TipoDocumento, NumeroDocumento, Telefono, Correo,
                Activo, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NumeroClientePorNegocio, @NegocioId, @NombreSolicitante, N'0', CONCAT(N'SOL', @Id), @Telefono, @Correo,
                1, SYSUTCDATETIME(), @Usuario
            );

            SET @ClienteId = SCOPE_IDENTITY();
        END;

        DECLARE @NumeroReservaPorNegocio INT;
        EXEC dbo.Sp_NegocioCorrelativos_ObtenerSiguiente @NegocioId, N'RESERVA', @Usuario, @NumeroReservaPorNegocio OUTPUT;

        INSERT INTO dbo.Reservas
        (
            NumeroPorNegocio, EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin,
            Estado, Total, Adelanto, Saldo, CodigoMoneda, CanalOrigen, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NumeroReservaPorNegocio, @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin,
            @EstadoReserva, @Total, @Adelanto, (@Total - @Adelanto), @CodigoMoneda, N'CLIENTE_WEB', SYSUTCDATETIME(), @Usuario
        );

        SET @ReservaId = SCOPE_IDENTITY();

        UPDATE dbo.SolicitudesReservaPublica
        SET Estado = 4,
            ReservaId = @ReservaId,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario,
            ComentarioGestion = N'Convertida a reserva'
        WHERE Id = @Id;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SOLICITUDES', @Accion = N'EDIT', @Entidad = N'SolicitudReservaPublica', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @ReservaId);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        DECLARE @MensajeNotificacion NVARCHAR(300);
        DECLARE @UrlNotificacion NVARCHAR(300);
        SET @MensajeNotificacion = N'Reserva R-' + RIGHT(N'000000' + CONVERT(NVARCHAR(20), @NumeroReservaPorNegocio), 6) + N' generada desde solicitud del cliente.';
        SET @UrlNotificacion = N'/Reservas?negocioId=' + CONVERT(NVARCHAR(20), @NegocioId);

        EXEC dbo.Sp_Notificaciones_Crear
            @NegocioId = @NegocioId,
            @Tipo = N'RESERVA_CLIENTE_WEB',
            @Titulo = N'Nueva reserva web',
            @Mensaje = @MensajeNotificacion,
            @Entidad = N'Reserva',
            @EntidadId = @ReservaId,
            @UrlDestino = @UrlNotificacion;

        COMMIT TRANSACTION;

        SELECT
            @ReservaId AS ReservaId,
            @NumeroReservaPorNegocio AS NumeroPorNegocio;
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
