SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/10/2026
-- Description:   Recupera en una consulta el detalle de las cabeceras del resultado exportado, conservando su orden.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_ExportarDetalle
    @idEmpresa CHAR(2),
    @idSucursal VARCHAR(8),
    @idDocumento CHAR(2),
    @Documentos XML
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT x.Documento.value('@orden', 'INT') AS OrdenCabecera,
               d.item, d.idProducto,
               CAST(CAST(d.GlsProducto AS VARCHAR(500)) AS VARCHAR(8000)) AS GlsProducto,
               d.GlsMarca, d.GlsUM,
               CAST(d.Cantidad AS NUMERIC(12,2)) AS Cantidad,
               CAST(d.Cantidad2 AS NUMERIC(12,2)) AS Cantidad2,
               CAST(d.VVUnit AS NUMERIC(12,2)) AS PVUnit,
               d.PorDcto,
               CAST(d.TotalVVNeto AS NUMERIC(12,2)) AS TotalPVNeto,
               ISNULL(d.CantidadImp, 0) AS CantidadImp,
               CASE WHEN d.item IS NULL THEN NULL
                    WHEN ISNULL(d.Comprado, '0') = '1' THEN 1 ELSE 0 END AS Comprado,
               ISNULL(d.GlsProveedor, '') AS GlsProveedor
        FROM @Documentos.nodes('/documentos/documento') x(Documento)
        LEFT JOIN dbo.docventasdet d
            ON d.idEmpresa = @idEmpresa
           AND d.idSucursal = @idSucursal
           AND d.idDocumento = @idDocumento
           AND d.idSerie = x.Documento.value('@serie', 'CHAR(4)')
           AND d.idDocVentas = x.Documento.value('@numero', 'CHAR(8)')
        ORDER BY x.Documento.value('@orden', 'INT'), d.item;
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
GO
