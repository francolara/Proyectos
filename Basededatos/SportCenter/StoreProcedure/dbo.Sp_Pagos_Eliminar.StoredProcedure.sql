USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 07_Reservas_Pagos_Reglas.sql (linea 340)
-- Firma: Codex - 09/04/2026 | Al eliminar pago: si reserva queda sin pagos cambia a Cancelada; si mantiene pagos cambia a Confirmada y recalcula adelanto/saldo.
-- Firma: Codex - 08/05/2026 | Bloquea eliminacion de pagos cuando la reserva ya tiene comprobante activo generado.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Eliminar]
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @ReservaId INT;

        SELECT @ReservaId = p.ReservaId
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE p.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @ReservaId IS NULL
            RAISERROR('No se encontro el pago para eliminar en el negocio.', 16, 1);

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
        WHERE Id = @Id;

        DECLARE @PagadoRestante DECIMAL(10,2);
        DECLARE @CantidadPagosRestante INT;

        SELECT
            @PagadoRestante = COALESCE(SUM(p.Monto), 0),
            @CantidadPagosRestante = COUNT(1)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        UPDATE dbo.Reservas
        SET Adelanto = @PagadoRestante,
            Saldo = (Total - @PagadoRestante),
            Estado = CASE WHEN @CantidadPagosRestante = 0 THEN 5 ELSE 2 END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @ReservaId;

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'PAGOS',
            @Accion = N'DELETE',
            @Entidad = N'Pago',
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
