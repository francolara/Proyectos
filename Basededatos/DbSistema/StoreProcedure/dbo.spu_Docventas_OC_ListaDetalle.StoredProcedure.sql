-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Obtiene el detalle resumido mostrado debajo del listado de ordenes de compra.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_OC_ListaDetalle
    @IdEmpresa CHAR(2),
    @IdSucursal VARCHAR(8),
    @IdDocumento CHAR(2),
    @IdSerie CHAR(4),
    @IdDocVentas CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            D.Item,
            D.IdProducto,
            CAST(D.GlsProducto AS VARCHAR(500)) AS GlsProducto,
            D.GlsMarca,
            D.GlsUM,
            CAST(D.Cantidad AS DECIMAL(10, 2)) AS Cantidad,
            CAST(D.VVUnit AS DECIMAL(10, 2)) AS PVUnit,
            CAST(D.PorDcto AS DECIMAL(10, 2)) AS PorDcto,
            CAST(D.TotalVVNeto AS DECIMAL(10, 2)) AS TotalPVNeto
        FROM dbo.DocVentasDet AS D
        WHERE D.IdEmpresa = @IdEmpresa
          AND D.IdSucursal = @IdSucursal
          AND D.IdDocumento = @IdDocumento
          AND D.IdSerie = @IdSerie
          AND D.IdDocVentas = @IdDocVentas
        ORDER BY D.Item;
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
