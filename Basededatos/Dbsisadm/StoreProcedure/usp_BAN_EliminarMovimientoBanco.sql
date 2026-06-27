-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Elimina movimientos de caja y bancos junto con su detalle.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Restaura el saldo de compras y ventas enlazadas antes de eliminar el movimiento bancario.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Elimina tambien el asiento contable vinculado al movimiento bancario.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Bloquea la eliminacion individual de movimientos que pertenecen a una transferencia entre cuentas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Restaura tambien el saldo de documentos de detraccion enlazados al eliminar un movimiento bancario.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_BAN_EliminarMovimientoBanco
    @IdMovimientoBanco INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdAsiento INT = NULL;
        DECLARE @IdTransferenciaCuenta UNIQUEIDENTIFIER = NULL;

        BEGIN TRANSACTION;

        SELECT
            @IdAsiento = m.IdAsiento,
            @IdTransferenciaCuenta = m.IdTransferenciaCuenta
        FROM dbo.BAN_MovimientoBanco AS m
        WHERE m.IdMovimientoBanco = @IdMovimientoBanco
          AND m.IdEmpresa = @IdEmpresa;

        IF @IdTransferenciaCuenta IS NOT NULL
        BEGIN
            RAISERROR('El movimiento pertenece a una transferencia entre cuentas. Elimine la transferencia completa desde su modulo.', 16, 1);
        END;

        ;WITH AplicacionesPrevias AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM dbo.BAN_MovimientoBancoDetalle AS d
            INNER JOIN dbo.BAN_MovimientoBanco AS m
                ON m.IdMovimientoBanco = d.IdMovimientoBanco
            WHERE d.IdMovimientoBanco = @IdMovimientoBanco
              AND m.IdEmpresa = @IdEmpresa
              AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE c
        SET c.Saldo = CASE
                          WHEN c.Saldo + a.ImporteAplicado > c.ImporteTotal THEN c.ImporteTotal
                          ELSE c.Saldo + a.ImporteAplicado
                      END
        FROM dbo.COM_Compra AS c
        INNER JOIN AplicacionesPrevias AS a
            ON a.ModuloOperacionComprobante = 'COM'
           AND a.IdRegistroComprobante = c.IdCompra
        WHERE c.IdEmpresa = @IdEmpresa;

        ;WITH AplicacionesPrevias AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM dbo.BAN_MovimientoBancoDetalle AS d
            INNER JOIN dbo.BAN_MovimientoBanco AS m
                ON m.IdMovimientoBanco = d.IdMovimientoBanco
            WHERE d.IdMovimientoBanco = @IdMovimientoBanco
              AND m.IdEmpresa = @IdEmpresa
              AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE cd
        SET cd.Saldo = CASE
                           WHEN cd.Saldo + a.ImporteAplicado > cd.ImporteDetraccion THEN cd.ImporteDetraccion
                           ELSE cd.Saldo + a.ImporteAplicado
                       END
        FROM dbo.COM_CompraDetraccion AS cd
        INNER JOIN AplicacionesPrevias AS a
            ON a.ModuloOperacionComprobante = 'DET'
           AND a.IdRegistroComprobante = cd.IdCompraDetraccion
        WHERE cd.IdEmpresa = @IdEmpresa;

        ;WITH AplicacionesPrevias AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM dbo.BAN_MovimientoBancoDetalle AS d
            INNER JOIN dbo.BAN_MovimientoBanco AS m
                ON m.IdMovimientoBanco = d.IdMovimientoBanco
            WHERE d.IdMovimientoBanco = @IdMovimientoBanco
              AND m.IdEmpresa = @IdEmpresa
              AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE v
        SET v.Saldo = CASE
                          WHEN v.Saldo + a.ImporteAplicado > v.ImporteTotal THEN v.ImporteTotal
                          ELSE v.Saldo + a.ImporteAplicado
                      END
        FROM dbo.VEN_Venta AS v
        INNER JOIN AplicacionesPrevias AS a
            ON a.ModuloOperacionComprobante = 'VEN'
           AND a.IdRegistroComprobante = v.IdVenta
        WHERE v.IdEmpresa = @IdEmpresa;

        DELETE d
        FROM dbo.BAN_MovimientoBancoDetalle AS d
        INNER JOIN dbo.BAN_MovimientoBanco AS m
            ON m.IdMovimientoBanco = d.IdMovimientoBanco
        WHERE d.IdMovimientoBanco = @IdMovimientoBanco
          AND m.IdEmpresa = @IdEmpresa;

        DELETE FROM dbo.BAN_MovimientoBanco
        WHERE IdMovimientoBanco = @IdMovimientoBanco
          AND IdEmpresa = @IdEmpresa;

        IF @@ROWCOUNT = 0
        BEGIN
            RAISERROR('El movimiento no existe para la empresa activa.', 16, 1);
        END;

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
