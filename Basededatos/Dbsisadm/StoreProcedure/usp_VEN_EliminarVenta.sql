-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Elimina una provision de venta y su asiento automatico vinculado desde el modulo de ventas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Impide eliminar ventas con cobros aplicados; solo permite eliminar si el comprobante sigue pendiente.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_VEN_EliminarVenta
    @IdVenta INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdAsiento INT
        DECLARE @ImporteTotal DECIMAL(18, 2)
        DECLARE @Saldo DECIMAL(18, 2)

        SELECT
            @IdAsiento = v.IdAsiento,
            @ImporteTotal = v.ImporteTotal,
            @Saldo = v.Saldo
        FROM dbo.VEN_Venta AS v
        WHERE v.IdVenta = @IdVenta
          AND v.IdEmpresa = @IdEmpresa;

        IF @IdAsiento IS NULL AND NOT EXISTS
        (
            SELECT 1
            FROM dbo.VEN_Venta AS v
            WHERE v.IdVenta = @IdVenta
              AND v.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La venta indicada no existe para la empresa activa.', 16, 1);
        END;

        IF ISNULL(@Saldo, 0) < ISNULL(@ImporteTotal, 0)
        BEGIN
            RAISERROR(N'La venta no puede eliminarse porque ya tiene cobros aplicados. Primero elimine el recibo o movimiento bancario relacionado.', 16, 1);
        END;

        BEGIN TRAN;

        DELETE FROM dbo.VEN_VentaDetalle
        WHERE IdVenta = @IdVenta;

        DELETE FROM dbo.VEN_Venta
        WHERE IdVenta = @IdVenta
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
