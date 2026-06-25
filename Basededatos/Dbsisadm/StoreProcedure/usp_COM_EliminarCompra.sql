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

CREATE OR ALTER PROCEDURE dbo.usp_COM_EliminarCompra
    @IdCompra INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdAsiento INT
        DECLARE @ImporteTotal DECIMAL(18, 2)
        DECLARE @Saldo DECIMAL(18, 2)

        SELECT
            @IdAsiento = c.IdAsiento,
            @ImporteTotal = c.ImporteTotal,
            @Saldo = c.Saldo
        FROM dbo.COM_Compra AS c
        WHERE c.IdCompra = @IdCompra
          AND c.IdEmpresa = @IdEmpresa;

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

        BEGIN TRAN;

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
