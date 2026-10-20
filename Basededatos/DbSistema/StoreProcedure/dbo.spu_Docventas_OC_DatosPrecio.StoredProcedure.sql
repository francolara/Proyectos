-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Obtiene el ultimo precio de compra de un producto para un proveedor.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_OC_DatosPrecio
    @IdEmpresa CHAR(2),
    @IdProveedor VARCHAR(20),
    @IdProducto VARCHAR(20),
    @TipoCambio DECIMAL(18, 6)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT TOP (1)
            U.AbreUM AS GlsUM,
            CASE
                WHEN PP.IdMoneda = 'PEN' THEN PP.ValorVenta
                ELSE PP.ValorVenta * @TipoCambio
            END AS VVUnit,
            CAST(1 AS DECIMAL(14, 6)) AS Factor
        FROM dbo.ProductosProveedor AS PP
        INNER JOIN dbo.Productos AS P
            ON P.IdEmpresa = PP.IdEmpresa
           AND P.IdProducto = PP.IdProducto
        INNER JOIN dbo.UnidadMedida AS U
            ON U.IdUM = P.IdUMCompra
        WHERE PP.IdEmpresa = @IdEmpresa
          AND PP.IdProveedor = @IdProveedor
          AND PP.IdProducto = @IdProducto
        ORDER BY PP.Fecha DESC;
    END TRY
    BEGIN CATCH
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
