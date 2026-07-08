-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Elimina una provision de compra y su asiento automatico vinculado desde el modulo de compras.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Impide eliminar compras con pagos aplicados; solo permite eliminar si el comprobante sigue pendiente.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Elimina tambien los documentos y asientos de detraccion y percepcion, validando que ambos saldos sigan pendientes.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Impide eliminar compras cuando la percepcion vinculada ya tenga pagos aplicados y depura su asiento asociado al eliminar.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Incluye la validacion y depuracion del pendiente COM_CompraRetencion al eliminar compras de recibos por honorarios.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Corrige la eliminacion de compras con detraccion para validar el saldo neto exigible del comprobante principal y no tratar la detraccion como pago aplicado.

CREATE OR ALTER PROCEDURE dbo.usp_COM_EliminarCompra
    @IdCompra INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdAsiento INT
        DECLARE @IdAsientoDetraccion INT
        DECLARE @IdCompraDetraccion INT
        DECLARE @IdAsientoPercepcion INT
        DECLARE @IdCompraPercepcion INT
        DECLARE @IdCompraRetencion INT
        DECLARE @ImporteTotal DECIMAL(18, 2)
        DECLARE @Saldo DECIMAL(18, 2)
        DECLARE @SaldoCompraExigible DECIMAL(18, 2)
        DECLARE @Retencion DECIMAL(18, 2)
        DECLARE @SaldoRetencion DECIMAL(18, 2)
        DECLARE @ImporteDetraccion DECIMAL(18, 2)
        DECLARE @SaldoDetraccion DECIMAL(18, 2)
        DECLARE @ImportePercepcion DECIMAL(18, 2)
        DECLARE @SaldoPercepcion DECIMAL(18, 2)

        SELECT
            @IdAsiento = c.IdAsiento,
            @ImporteTotal = c.ImporteTotal,
            @Saldo = c.Saldo,
            @Retencion = c.Retencion,
            @ImporteDetraccion = c.ImporteDetraccion,
            @ImportePercepcion = c.ImportePercepcion
        FROM dbo.COM_Compra AS c
        WHERE c.IdCompra = @IdCompra
          AND c.IdEmpresa = @IdEmpresa;

        SET @SaldoCompraExigible = ISNULL(@ImporteTotal, 0) - ISNULL(@ImporteDetraccion, 0);
        IF @SaldoCompraExigible < 0
        BEGIN
            SET @SaldoCompraExigible = 0;
        END;

        SELECT
            @IdCompraRetencion = cr.IdCompraRetencion,
            @SaldoRetencion = cr.Saldo
        FROM dbo.COM_CompraRetencion AS cr
        WHERE cr.IdCompra = @IdCompra
          AND cr.IdEmpresa = @IdEmpresa;

        SELECT
            @IdCompraDetraccion = cd.IdCompraDetraccion,
            @IdAsientoDetraccion = cd.IdAsiento,
            @SaldoDetraccion = cd.Saldo
        FROM dbo.COM_CompraDetraccion AS cd
        WHERE cd.IdCompra = @IdCompra
          AND cd.IdEmpresa = @IdEmpresa;

        SELECT
            @IdCompraPercepcion = cp.IdCompraPercepcion,
            @IdAsientoPercepcion = cp.IdAsiento,
            @SaldoPercepcion = cp.Saldo
        FROM dbo.COM_CompraPercepcion AS cp
        WHERE cp.IdCompra = @IdCompra
          AND cp.IdEmpresa = @IdEmpresa;

        IF @IdAsiento IS NULL AND NOT EXISTS
        (
            SELECT 1
            FROM dbo.COM_Compra AS c
            WHERE c.IdCompra = @IdCompra
              AND c.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La compra indicada no existe para la empresa activa.', 16, 1);
        END;

        IF ISNULL(@Saldo, 0) < @SaldoCompraExigible
        BEGIN
            RAISERROR(N'La compra no puede eliminarse porque ya tiene pagos aplicados. Primero elimine el recibo o movimiento bancario relacionado.', 16, 1);
        END;

        IF @IdCompraRetencion IS NOT NULL
           AND ISNULL(@SaldoRetencion, 0) < ISNULL(@Retencion, 0)
        BEGIN
            RAISERROR(N'La compra no puede eliminarse porque la retencion de renta de 4ta ya tiene pagos aplicados. Primero elimine el pago relacionado.', 16, 1);
        END;

        IF @IdCompraDetraccion IS NOT NULL
           AND ISNULL(@SaldoDetraccion, 0) < ISNULL(@ImporteDetraccion, 0)
        BEGIN
            RAISERROR(N'La compra no puede eliminarse porque la detraccion ya tiene pagos aplicados. Primero elimine el pago de detraccion relacionado.', 16, 1);
        END;

        IF @IdCompraPercepcion IS NOT NULL
           AND ISNULL(@SaldoPercepcion, 0) < ISNULL(@ImportePercepcion, 0)
        BEGIN
            RAISERROR(N'La compra no puede eliminarse porque la percepcion ya tiene pagos aplicados. Primero elimine el pago de percepcion relacionado.', 16, 1);
        END;

        BEGIN TRAN;

        IF @IdCompraRetencion IS NOT NULL
        BEGIN
            DELETE FROM dbo.COM_CompraRetencion
            WHERE IdCompraRetencion = @IdCompraRetencion;
        END;

        IF @IdCompraDetraccion IS NOT NULL
        BEGIN
            DELETE FROM dbo.COM_CompraDetraccion
            WHERE IdCompraDetraccion = @IdCompraDetraccion;

            IF @IdAsientoDetraccion IS NOT NULL
            BEGIN
                DELETE FROM dbo.CON_AsientoDetalle
                WHERE IdAsiento = @IdAsientoDetraccion;

                DELETE FROM dbo.CON_Asiento
                WHERE IdAsiento = @IdAsientoDetraccion
                  AND IdEmpresa = @IdEmpresa;
            END;
        END;

        IF @IdCompraPercepcion IS NOT NULL
        BEGIN
            DELETE FROM dbo.COM_CompraPercepcion
            WHERE IdCompraPercepcion = @IdCompraPercepcion;

            IF @IdAsientoPercepcion IS NOT NULL
            BEGIN
                DELETE FROM dbo.CON_AsientoDetalle
                WHERE IdAsiento = @IdAsientoPercepcion;

                DELETE FROM dbo.CON_Asiento
                WHERE IdAsiento = @IdAsientoPercepcion
                  AND IdEmpresa = @IdEmpresa;
            END;
        END;

        DELETE FROM dbo.COM_CompraDetalle
        WHERE IdCompra = @IdCompra;

        DELETE FROM dbo.COM_Compra
        WHERE IdCompra = @IdCompra
          AND IdEmpresa = @IdEmpresa;

        IF @IdAsiento IS NOT NULL
        BEGIN
            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsiento;

            DELETE FROM dbo.CON_Asiento
            WHERE IdAsiento = @IdAsiento
              AND IdEmpresa = @IdEmpresa;
        END;

        COMMIT TRAN;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRAN;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
