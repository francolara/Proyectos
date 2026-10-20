-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista productos y precios asociados a un proveedor.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedores_Productos_Listar
    @IdEmpresa CHAR(2),
    @IdProveedor CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PP.Item,
            PP.IdProducto,
            P.GlsProducto,
            PP.ValorVenta,
            PP.IGVVenta,
            PP.PrecioVenta
        FROM dbo.ProductosProveedor AS PP
        INNER JOIN dbo.Productos AS P
            ON PP.IdEmpresa = P.IdEmpresa
           AND PP.IdProducto = P.IdProducto
        WHERE PP.IdEmpresa = @IdEmpresa
          AND PP.IdProveedor = @IdProveedor
        ORDER BY PP.Item;
    END TRY
    BEGIN CATCH
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
