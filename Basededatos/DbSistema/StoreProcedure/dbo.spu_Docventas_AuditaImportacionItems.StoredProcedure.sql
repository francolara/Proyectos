
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/10/2026
-- Description:   Lista importaciones historicas ambiguas cuando el documento origen repite producto y UM.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_AuditaImportacionItems
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10) = '',
    @IdDocumentoOrigen VARCHAR(10) = '',
    @IdSerieOrigen VARCHAR(10) = '',
    @IdDocVentasOrigen VARCHAR(20) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH DetallesOrigen AS
        (
            SELECT
                origen.IdEmpresa,
                origen.IdSucursal,
                origen.IdDocumento,
                origen.IdSerie,
                origen.IdDocVentas,
                origen.Item,
                origen.IdProducto,
                origen.IdUM,
                origen.Cantidad,
                origen.CantidadImp,
                origen.EstDocImportado,
                COUNT(1) OVER
                (
                    PARTITION BY
                        origen.IdEmpresa,
                        origen.IdSucursal,
                        origen.IdDocumento,
                        origen.IdSerie,
                        origen.IdDocVentas,
                        origen.IdProducto,
                        origen.IdUM
                ) AS RepeticionesProducto
            FROM dbo.DocVentasDet AS origen
            INNER JOIN dbo.DocVentas AS cabeceraOrigen
                ON cabeceraOrigen.IdEmpresa = origen.IdEmpresa
                AND cabeceraOrigen.IdSucursal = origen.IdSucursal
                AND cabeceraOrigen.IdDocumento = origen.IdDocumento
                AND cabeceraOrigen.IdSerie = origen.IdSerie
                AND cabeceraOrigen.IdDocVentas = origen.IdDocVentas
            WHERE origen.IdEmpresa = @IdEmpresa
                AND (@IdSucursal = '' OR origen.IdSucursal = @IdSucursal)
                AND (@IdDocumentoOrigen = '' OR origen.IdDocumento = @IdDocumentoOrigen)
                AND (@IdSerieOrigen = '' OR origen.IdSerie = @IdSerieOrigen)
                AND (@IdDocVentasOrigen = '' OR origen.IdDocVentas = @IdDocVentasOrigen)
                AND ISNULL(cabeceraOrigen.EstDocVentas, '') <> 'ANU'
        )
        SELECT
            origen.IdEmpresa,
            origen.IdSucursal,
            origen.IdDocumento AS IdDocumentoOrigen,
            origen.IdSerie AS IdSerieOrigen,
            origen.IdDocVentas AS IdDocVentasOrigen,
            origen.Item AS ItemOrigenPosible,
            origen.IdProducto,
            origen.IdUM,
            origen.Cantidad AS CantidadOrigen,
            origen.CantidadImp AS CantidadImportadaOrigen,
            origen.EstDocImportado AS EstadoImportacionOrigen,
            destino.IdDocumento AS IdDocumentoDestino,
            destino.IdSerie AS IdSerieDestino,
            destino.IdDocVentas AS IdDocVentasDestino,
            destino.Item AS ItemDestino,
            destino.ItemPro AS ItemOrigenRegistrado,
            destino.Cantidad AS CantidadDestino,
            cabeceraDestino.EstDocVentas AS EstadoDocumentoDestino,
            CASE
                WHEN destino.ItemPro = origen.Item THEN 1
                ELSE 0
            END AS CoincideItemRegistrado,
            CASE
                WHEN destino.ItemPro = origen.Item THEN 'REVISAR_ASIGNACION_HISTORICA'
                ELSE 'ITEM_ORIGEN_NO_COINCIDE'
            END AS Diagnostico
        FROM DetallesOrigen AS origen
        INNER JOIN dbo.DocVentasDet AS destino
            ON destino.IdEmpresa = origen.IdEmpresa
            AND destino.IdSucursal = origen.IdSucursal
            AND destino.IdDocumentoImp = origen.IdDocumento
            AND destino.IdSerieImp = origen.IdSerie
            AND destino.IdDocVentasImp = origen.IdDocVentas
            AND destino.IdProducto = origen.IdProducto
            AND destino.IdUM = origen.IdUM
        INNER JOIN dbo.DocVentas AS cabeceraDestino
            ON cabeceraDestino.IdEmpresa = destino.IdEmpresa
            AND cabeceraDestino.IdSucursal = destino.IdSucursal
            AND cabeceraDestino.IdDocumento = destino.IdDocumento
            AND cabeceraDestino.IdSerie = destino.IdSerie
            AND cabeceraDestino.IdDocVentas = destino.IdDocVentas
        WHERE origen.RepeticionesProducto > 1
        ORDER BY
            origen.IdDocumento,
            origen.IdSerie,
            origen.IdDocVentas,
            origen.IdProducto,
            origen.IdUM,
            destino.IdDocumento,
            destino.IdSerie,
            destino.IdDocVentas,
            destino.Item,
            origen.Item;
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
