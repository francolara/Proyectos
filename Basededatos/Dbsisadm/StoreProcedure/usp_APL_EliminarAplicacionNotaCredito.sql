-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Elimina una aplicacion de nota de credito restaurando saldos y eliminando su asiento vinculado.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Agrega la reversa operativa del modulo Aplicaciones para devolver saldo al comprobante y a la NC antes de borrar el asiento asociado.

CREATE OR ALTER PROCEDURE dbo.usp_APL_EliminarAplicacionNotaCredito
    @IdAplicacionNotaCredito INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @ModuloOperacion VARCHAR(10) = NULL;
        DECLARE @IdRegistroComprobante INT = NULL;
        DECLARE @IdRegistroNotaCredito INT = NULL;
        DECLARE @ImporteAplicado DECIMAL(18, 2) = 0;
        DECLARE @IdAsiento INT = NULL;

        SELECT
            @ModuloOperacion = a.ModuloOperacion,
            @IdRegistroComprobante = a.IdRegistroComprobante,
            @IdRegistroNotaCredito = a.IdRegistroNotaCredito,
            @ImporteAplicado = a.ImporteAplicado,
            @IdAsiento = a.IdAsiento
        FROM dbo.CON_AplicacionNotaCredito AS a
        WHERE a.IdAplicacionNotaCredito = @IdAplicacionNotaCredito
          AND a.IdEmpresa = @IdEmpresa
          AND a.Activo = 1;

        IF @ModuloOperacion IS NULL
        BEGIN
            RAISERROR(N'La aplicacion seleccionada no existe para la empresa activa.', 16, 1);
        END;

        BEGIN TRANSACTION;

        IF @ModuloOperacion = 'VEN'
        BEGIN
            UPDATE dbo.VEN_Venta
            SET Saldo = CASE
                            WHEN Saldo + @ImporteAplicado > ImporteTotal THEN ImporteTotal
                            ELSE Saldo + @ImporteAplicado
                        END
            WHERE IdEmpresa = @IdEmpresa
              AND IdVenta IN (@IdRegistroComprobante, @IdRegistroNotaCredito);
        END
        ELSE
        BEGIN
            UPDATE dbo.COM_Compra
            SET Saldo = CASE
                            WHEN Saldo + @ImporteAplicado > ImporteTotal THEN ImporteTotal
                            ELSE Saldo + @ImporteAplicado
                        END
            WHERE IdEmpresa = @IdEmpresa
              AND IdCompra IN (@IdRegistroComprobante, @IdRegistroNotaCredito);
        END;

        DELETE FROM dbo.CON_AplicacionNotaCredito
        WHERE IdAplicacionNotaCredito = @IdAplicacionNotaCredito
          AND IdEmpresa = @IdEmpresa;

        IF @IdAsiento IS NOT NULL
        BEGIN
            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsiento;

            DELETE FROM dbo.CON_Asiento
            WHERE IdAsiento = @IdAsiento
              AND IdEmpresa = @IdEmpresa;
        END;

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);

    END CATCH

END
