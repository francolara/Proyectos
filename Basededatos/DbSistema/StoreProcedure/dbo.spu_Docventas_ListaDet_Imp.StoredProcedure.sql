
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/10/2026
-- Description:   Lista el detalle disponible y conserva el item del documento origen inmediato.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_ListaDet_Imp
    @IdEmpresa CHAR(2),
    @IdSucursal VARCHAR(8),
    @IdDocumento CHAR(2),
    @IdDocVentas CHAR(8),
    @IdSerie CHAR(4)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            dvd.Item,
            dvd.Item AS ItemOrigen,
            dvd.IdProducto,
            dvd.IdCodFabricante,
            dvd.GlsProducto,
            dvd.IdMarca,
            dvd.GlsMarca,
            dvd.IdUM,
            dvd.GlsUM,
            dvd.Factor,
            dvd.Afecto,
            dvd.Cantidad - ISNULL(dvd.CantidadImp, 0) AS Cantidad,
            dvd.VVUnit,
            dvd.IGVUnit,
            dvd.PVUnit,
            dvd.TotalVVBruto,
            dvd.TotalPVBruto,
            dvd.PorDcto,
            dvd.DctoVV,
            dvd.DctoPV,
            dvd.TotalVVNeto,
            dvd.TotalIGVNeto,
            dvd.TotalPVNeto,
            dvd.IdTipoProducto,
            dvd.IdMoneda,
            dvd.NumLote,
            dvd.FecVencProd,
            dvd.VVUnitLista,
            dvd.PVUnitLista,
            dvd.VVUnitNeto,
            dvd.PVUnitNeto,
            dvd.Cantidad2,
            dvd.CodigoRapido,
            dvd.IdTallaPeso,
            dvd.ItemPro,
            dvd.IvapUnit,
            dvd.TotalIvapNeto,
            dvd.IdDocVentas,
            dvd.IdSerie,
            dvd.IdLote,
            dvd.GlsProveedor,
            dvd.CostoS,
            dvd.CostoSInc,
            dvd.CostoD,
            dvd.Margen
        FROM dbo.DocVentasDet AS dvd
        WHERE dvd.IdEmpresa = @IdEmpresa
            AND dvd.IdSucursal = @IdSucursal
            AND dvd.IdDocumento = @IdDocumento
            AND dvd.IdDocVentas = @IdDocVentas
            AND dvd.IdSerie = @IdSerie
            AND ISNULL(dvd.EstDocImportado, 'N') <> 'S'
        ORDER BY dvd.Item;
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
GO
