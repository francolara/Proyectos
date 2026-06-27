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
-- Description:   Elimina tambien el documento y asiento de detraccion, validando que ambos saldos sigan pendientes.
-- =============================================

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
        DECLARE @ImporteTotal DECIMAL(18, 2)
        DECLARE @Saldo DECIMAL(18, 2)
        DECLARE @ImporteDetraccion DECIMAL(18, 2)
        DECLARE @SaldoDetraccion DECIMAL(18, 2)

        SELECT
            @IdAsiento = c.IdAsiento,
            @ImporteTotal = c.ImporteTotal,
            @Saldo = c.Saldo,
            @ImporteDetraccion = c.ImporteDetraccion
        FROM dbo.COM_Compra AS c
        WHERE c.IdCompra = @IdCompra
          AND c.IdEmpresa = @IdEmpresa;

        SELECT
            @IdCompraDetraccion = cd.IdCompraDetraccion,
            @IdAsientoDetraccion = cd.IdAsiento,
            @SaldoDetraccion = cd.Saldo
        FROM dbo.COM_CompraDetraccion AS cd
        WHERE cd.IdCompra = @IdCompra
          AND cd.IdEmpresa = @IdEmpresa;

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

        IF ISNULL(@Saldo, 0) < ISNULL(@ImporteTotal, 0)
        BEGIN
            RAISERROR(N'La compra no puede eliminarse porque ya tiene pagos aplicados. Primero elimine el recibo o movimiento bancario relacionado.', 16, 1);
        END;

        IF @IdCompraDetraccion IS NOT NULL
           AND ISNULL(@SaldoDetraccion, 0) < ISNULL(@ImporteDetraccion, 0)
        BEGIN
            RAISERROR(N'La compra no puede eliminarse porque la detraccion ya tiene pagos aplicados. Primero elimine el pago de detraccion relacionado.', 16, 1);
        END;

        BEGIN TRAN;

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
