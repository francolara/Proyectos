USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 35_Maestros_FormasPago.sql (linea 186)
-- Firma: Codex - 09/04/2026 | Limita a maximo 2 pagos por reserva, valida politica de confirmacion del negocio, exige que el 2do pago sea exactamente el saldo restante y ajusta estado automatico segun pago acumulado.
-- Firma: FRANCO LARA - 16/07/2026 | Permite pagos parciales sin limite de cantidad y bloquea reservas pagadas o montos mayores al saldo pendiente.
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
        DECLARE @SaldoPendiente DECIMAL(10,2);
        DECLARE @PoliticaConfirmacionPago TINYINT = 0;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL;
        DECLARE @MontoMinimoAdelanto DECIMAL(10,2) = NULL;

        SELECT
            @TotalReserva = r.Total,
            @PoliticaConfirmacionPago = ISNULL(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @TotalReserva IS NULL
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        SELECT
            @PagadoActual = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        SET @SaldoPendiente = @TotalReserva - @PagadoActual;

        IF @SaldoPendiente <= 0
            RAISERROR('La reserva ya esta pagada al 100%. No se pueden registrar mas pagos.', 16, 1);

        IF @Monto > @SaldoPendiente
            RAISERROR('El pago excede el saldo pendiente de la reserva.', 16, 1);

        SET @NuevoPagado = @PagadoActual + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        IF @PoliticaConfirmacionPago = 2 AND @NuevoPagado < @TotalReserva
            RAISERROR('La configuracion del negocio exige pago total (100%) para confirmar la reserva.', 16, 1);

        IF @PoliticaConfirmacionPago = 1
        BEGIN
            SET @PorcentajeAdelantoMinimo = ISNULL(@PorcentajeAdelantoMinimo, 0);
            IF @PorcentajeAdelantoMinimo > 0
            BEGIN
                SET @MontoMinimoAdelanto = ROUND((@TotalReserva * @PorcentajeAdelantoMinimo) / 100.0, 2);
                IF @NuevoPagado < @MontoMinimoAdelanto AND @NuevoPagado < @TotalReserva
                    RAISERROR('El pago acumulado no alcanza el adelanto minimo configurado para confirmar la reserva.', 16, 1);
            END
        END

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
                        WHEN @NuevoPagado > 0 THEN 2
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
