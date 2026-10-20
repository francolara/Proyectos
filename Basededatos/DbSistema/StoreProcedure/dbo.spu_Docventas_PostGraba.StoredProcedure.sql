-- =============================================
-- Author:        FRANCO LARA
-- Create date:   14/10/2026
-- Description:   Consolida observaciones y peso total despues de grabar un documento de venta.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_PostGraba
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10),
    @IdDocumento VARCHAR(10),
    @IdSerie VARCHAR(20),
    @IdDocVentas VARCHAR(20),
    @ObsDocVentas VARCHAR(MAX),
    @ObsDocVentas2 VARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        UPDATE dbo.DocVentas
        SET ObsDocVentas = @ObsDocVentas,
            ObsDocVentas2 = @ObsDocVentas2,
            TotalTallaPeso = (
                SELECT ISNULL(SUM(Detalle.IdTallaPeso), 0)
                FROM dbo.DocVentasDet AS Detalle
                WHERE Detalle.IdEmpresa = @IdEmpresa
                  AND Detalle.IdSucursal = @IdSucursal
                  AND Detalle.IdDocumento = @IdDocumento
                  AND Detalle.IdSerie = @IdSerie
                  AND Detalle.IdDocVentas = @IdDocVentas
            )
        WHERE IdEmpresa = @IdEmpresa
          AND IdSucursal = @IdSucursal
          AND IdDocumento = @IdDocumento
          AND IdSerie = @IdSerie
          AND IdDocVentas = @IdDocVentas;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(),
               @ErrorSeverity = ERROR_SEVERITY(),
               @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
