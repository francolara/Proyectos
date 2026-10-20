
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/10/2026
-- Description:   Recalcula por item origen y mantiene compatibilidad historica por producto/UM sin duplicados.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_RecalculaImportacionDocumento
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10),
    @IdDocumentoOrigen VARCHAR(10),
    @IdSerieOrigen VARCHAR(10),
    @IdDocVentasOrigen VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH DetallesOrigen AS
        (
            SELECT
                origen.Item,
                origen.IdProducto,
                origen.IdUM,
                COUNT(1) OVER
                (
                    PARTITION BY origen.IdProducto, origen.IdUM
                ) AS RepeticionesProducto
            FROM dbo.DocVentasDet AS origen
            WHERE origen.IdEmpresa = @IdEmpresa
                AND origen.IdSucursal = @IdSucursal
                AND origen.IdDocumento = @IdDocumentoOrigen
                AND origen.IdSerie = @IdSerieOrigen
                AND origen.IdDocVentas = @IdDocVentasOrigen
        ),
        ImportacionesVigentes AS
        (
            SELECT
                destino.ItemPro,
                destino.IdProducto,
                destino.IdUM,
                ISNULL(destino.Cantidad, 0) AS Cantidad
            FROM dbo.DocVentasDet AS destino
            INNER JOIN dbo.DocVentas AS cabeceraDestino
                ON cabeceraDestino.IdEmpresa = destino.IdEmpresa
                AND cabeceraDestino.IdSucursal = destino.IdSucursal
                AND cabeceraDestino.IdDocumento = destino.IdDocumento
                AND cabeceraDestino.IdSerie = destino.IdSerie
                AND cabeceraDestino.IdDocVentas = destino.IdDocVentas
            WHERE destino.IdEmpresa = @IdEmpresa
                AND destino.IdSucursal = @IdSucursal
                AND destino.IdDocumentoImp = @IdDocumentoOrigen
                AND destino.IdSerieImp = @IdSerieOrigen
                AND destino.IdDocVentasImp = @IdDocVentasOrigen
                AND ISNULL(cabeceraDestino.EstDocVentas, '') <> 'ANU'
        ),
        CantidadesPorItem AS
        (
            SELECT
                importacion.ItemPro AS ItemOrigen,
                SUM(importacion.Cantidad) AS CantidadImp
            FROM ImportacionesVigentes AS importacion
            WHERE ISNULL(importacion.ItemPro, 0) > 0
            GROUP BY importacion.ItemPro
        ),
        CantidadesPorProducto AS
        (
            SELECT
                importacion.IdProducto,
                importacion.IdUM,
                SUM(importacion.Cantidad) AS CantidadImp
            FROM ImportacionesVigentes AS importacion
            GROUP BY
                importacion.IdProducto,
                importacion.IdUM
        )
        UPDATE origen
        SET
            origen.CantidadImp = CASE
                WHEN detalleOrigen.RepeticionesProducto = 1
                    THEN ISNULL(importadoProducto.CantidadImp, 0)
                ELSE ISNULL(importadoItem.CantidadImp, 0)
            END,
            origen.EstDocImportado = CASE
                WHEN
                    CASE
                        WHEN detalleOrigen.RepeticionesProducto = 1
                            THEN ISNULL(importadoProducto.CantidadImp, 0)
                        ELSE ISNULL(importadoItem.CantidadImp, 0)
                    END >= ISNULL(origen.Cantidad, 0) THEN 'S'
                ELSE 'N'
            END
        FROM dbo.DocVentasDet AS origen
        INNER JOIN DetallesOrigen AS detalleOrigen
            ON detalleOrigen.Item = origen.Item
        LEFT JOIN CantidadesPorItem AS importadoItem
            ON importadoItem.ItemOrigen = origen.Item
        LEFT JOIN CantidadesPorProducto AS importadoProducto
            ON importadoProducto.IdProducto = origen.IdProducto
            AND importadoProducto.IdUM = origen.IdUM
        WHERE origen.IdEmpresa = @IdEmpresa
            AND origen.IdSucursal = @IdSucursal
            AND origen.IdDocumento = @IdDocumentoOrigen
            AND origen.IdSerie = @IdSerieOrigen
            AND origen.IdDocVentas = @IdDocVentasOrigen;

        UPDATE cabeceraOrigen
        SET cabeceraOrigen.EstDocImportado = CASE
            WHEN EXISTS
            (
                SELECT 1
                FROM dbo.DocVentasDet AS detalleOrigen
                WHERE detalleOrigen.IdEmpresa = @IdEmpresa
                    AND detalleOrigen.IdSucursal = @IdSucursal
                    AND detalleOrigen.IdDocumento = @IdDocumentoOrigen
                    AND detalleOrigen.IdSerie = @IdSerieOrigen
                    AND detalleOrigen.IdDocVentas = @IdDocVentasOrigen
            )
            AND NOT EXISTS
            (
                SELECT 1
                FROM dbo.DocVentasDet AS detalleOrigen
                WHERE detalleOrigen.IdEmpresa = @IdEmpresa
                    AND detalleOrigen.IdSucursal = @IdSucursal
                    AND detalleOrigen.IdDocumento = @IdDocumentoOrigen
                    AND detalleOrigen.IdSerie = @IdSerieOrigen
                    AND detalleOrigen.IdDocVentas = @IdDocVentasOrigen
                    AND ISNULL(detalleOrigen.EstDocImportado, '') <> 'S'
            ) THEN 'S'
            ELSE 'N'
        END
        FROM dbo.DocVentas AS cabeceraOrigen
        WHERE cabeceraOrigen.IdEmpresa = @IdEmpresa
            AND cabeceraOrigen.IdSucursal = @IdSucursal
            AND cabeceraOrigen.IdDocumento = @IdDocumentoOrigen
            AND cabeceraOrigen.IdSerie = @IdSerieOrigen
            AND cabeceraOrigen.IdDocVentas = @IdDocVentasOrigen;
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
