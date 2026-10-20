
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/10/2026
-- Description:   Valida items importados e informa tipo, serie, numero y cantidad de cada documento destino vigente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_ValidaEliminaItemImportado
    @IdEmpresa VARCHAR(10),
    @IdSucursal VARCHAR(10),
    @IdDocumentoOrigen VARCHAR(10),
    @IdSerieOrigen VARCHAR(10),
    @IdDocVentasOrigen VARCHAR(20),
    @ItemOrigen INT = NULL
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
                destino.IdDocumento AS IdDocumentoDestino,
                destino.IdSerie AS IdSerieDestino,
                destino.IdDocVentas AS IdDocVentasDestino,
                ISNULL(documentoDestino.GlsDocumento, destino.IdDocumento) AS TipoDocumentoDestino,
                ISNULL(destino.Cantidad, 0) AS Cantidad
            FROM dbo.DocVentasDet AS destino
            INNER JOIN dbo.DocVentas AS cabeceraDestino
                ON cabeceraDestino.IdEmpresa = destino.IdEmpresa
                AND cabeceraDestino.IdSucursal = destino.IdSucursal
                AND cabeceraDestino.IdDocumento = destino.IdDocumento
                AND cabeceraDestino.IdSerie = destino.IdSerie
                AND cabeceraDestino.IdDocVentas = destino.IdDocVentas
            LEFT JOIN dbo.Documentos AS documentoDestino
                ON documentoDestino.IdDocumento = destino.IdDocumento
            WHERE destino.IdEmpresa = @IdEmpresa
                AND destino.IdSucursal = @IdSucursal
                AND destino.IdDocumentoImp = @IdDocumentoOrigen
                AND destino.IdSerieImp = @IdSerieOrigen
                AND destino.IdDocVentasImp = @IdDocVentasOrigen
                AND ISNULL(cabeceraDestino.EstDocVentas, '') <> 'ANU'
        )
        SELECT
            origen.Item,
            origen.IdProducto,
            origen.IdUM,
            CASE
                WHEN calculo.CantidadImportada > 0 THEN 1
                ELSE 0
            END AS ItemImportado,
            calculo.CantidadImportada,
            ISNULL(calculo.DetalleImportacion, '') AS DetalleImportacion
        FROM DetallesOrigen AS origen
        OUTER APPLY
        (
            SELECT
                ISNULL(SUM(documentoImportado.Cantidad), 0) AS CantidadImportada,
                STRING_AGG
                (
                    CONVERT
                    (
                        NVARCHAR(MAX),
                        documentoImportado.TipoDocumentoDestino + ' ' +
                        documentoImportado.IdSerieDestino + '-' +
                        documentoImportado.IdDocVentasDestino +
                        ', Cantidad: ' +
                        CONVERT(VARCHAR(30), CAST(documentoImportado.Cantidad AS DECIMAL(18, 2)))
                    ),
                    CHAR(13) + CHAR(10)
                ) WITHIN GROUP
                (
                    ORDER BY
                        documentoImportado.TipoDocumentoDestino,
                        documentoImportado.IdSerieDestino,
                        documentoImportado.IdDocVentasDestino
                ) AS DetalleImportacion
            FROM
            (
                SELECT
                    importacion.IdDocumentoDestino,
                    importacion.IdSerieDestino,
                    importacion.IdDocVentasDestino,
                    importacion.TipoDocumentoDestino,
                    SUM(ABS(importacion.Cantidad)) AS Cantidad
                FROM ImportacionesVigentes AS importacion
                WHERE
                    (
                        origen.RepeticionesProducto = 1
                        AND importacion.IdProducto = origen.IdProducto
                        AND importacion.IdUM = origen.IdUM
                    )
                    OR
                    (
                        origen.RepeticionesProducto > 1
                        AND importacion.ItemPro = origen.Item
                    )
                GROUP BY
                    importacion.IdDocumentoDestino,
                    importacion.IdSerieDestino,
                    importacion.IdDocVentasDestino,
                    importacion.TipoDocumentoDestino
            ) AS documentoImportado
        ) AS calculo
        WHERE @ItemOrigen IS NULL OR origen.Item = @ItemOrigen
        ORDER BY origen.Item;
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
