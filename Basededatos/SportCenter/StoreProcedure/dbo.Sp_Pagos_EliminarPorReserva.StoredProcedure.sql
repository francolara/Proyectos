USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 09/04/2026 | Elimina todos los pagos de una reserva y deja la reserva en estado Cancelada.
-- Firma: Codex - 08/05/2026 | Bloquea eliminacion masiva de pagos cuando la reserva ya tiene comprobante activo generado.
CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_EliminarPorReserva
    @NegocioId INT,
    @ReservaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE r.Id = @ReservaId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('No se encontro la reserva para eliminar pagos en el negocio.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM dbo.ComprobantesElectronicos ce
            WHERE ce.NegocioId = @NegocioId
              AND ce.ReservaId = @ReservaId
              AND ce.Estado IN(1,2,3)
              AND ce.ComprobanteReferenciaId IS NULL
        )
            RAISERROR('No se puede eliminar pagos porque la reserva ya tiene un comprobante generado.', 16, 1);

        BEGIN TRANSACTION;

        DELETE FROM dbo.Pagos
        WHERE ReservaId = @ReservaId;

        UPDATE dbo.Reservas
        SET Adelanto = 0,
            Saldo = Total,
            Estado = 5,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @ReservaId;

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @ReservaId);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'PAGOS',
            @Accion = N'DELETE',
            @Entidad = N'ReservaPago',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

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
