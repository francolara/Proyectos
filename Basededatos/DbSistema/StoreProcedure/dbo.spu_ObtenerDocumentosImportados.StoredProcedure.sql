
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/10/2026
-- Description:   Obtiene todos los documentos origen referenciados por un documento destino.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_ObtenerDocumentosImportados
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10),
    @IdDocumentoDestino VARCHAR(10),
    @IdSerieDestino VARCHAR(10),
    @IdDocVentasDestino VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT DISTINCT
            origen.IdEmpresa,
            origen.IdSucursal,
            dvd.IdDocumentoImp,
            dvd.IdSerieImp,
            dvd.IdDocVentasImp
        FROM dbo.DocVentasDet AS dvd
        INNER JOIN dbo.DocVentas AS origen
            ON origen.IdEmpresa = dvd.IdEmpresa
            AND origen.IdSucursal = dvd.IdSucursal
            AND origen.IdDocumento = dvd.IdDocumentoImp
            AND origen.IdSerie = dvd.IdSerieImp
            AND origen.IdDocVentas = dvd.IdDocVentasImp
        WHERE dvd.IdEmpresa = @IdEmpresa
            AND dvd.IdSucursal = @IdSucursal
            AND dvd.IdDocumento = @IdDocumentoDestino
            AND dvd.IdSerie = @IdSerieDestino
            AND dvd.IdDocVentas = @IdDocVentasDestino
            AND ISNULL(dvd.IdDocumentoImp, '') <> ''
            AND ISNULL(dvd.IdSerieImp, '') <> ''
            AND ISNULL(dvd.IdDocVentasImp, '') <> '';
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
