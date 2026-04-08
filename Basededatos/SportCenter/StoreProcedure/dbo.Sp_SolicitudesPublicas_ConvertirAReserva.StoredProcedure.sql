USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_SolicitudesPublicas_ConvertirAReserva para tipo de documento SUNAT por defecto.
-- Firma: Codex - 06/04/2026 | Si la solicitud se convierte como Confirmada, valida politica de pago del negocio.
-- Firma: Codex - 06/04/2026 | Se elimina dependencia de NegocioClientes y se usa Clientes.NegocioId.
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

        SELECT
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId;

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
            INSERT INTO dbo.Clientes
            (
                NegocioId, NombresORazonSocial, TipoDocumento, NumeroDocumento, Telefono, Correo,
                Activo, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, @NombreSolicitante, N'0', CONCAT(N'SOL', @Id), @Telefono, @Correo,
                1, SYSUTCDATETIME(), @Usuario
            );

            SET @ClienteId = SCOPE_IDENTITY();
        END;

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin,
            Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin,
            @EstadoReserva, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario
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

        COMMIT TRANSACTION;

        SELECT @ReservaId;
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
