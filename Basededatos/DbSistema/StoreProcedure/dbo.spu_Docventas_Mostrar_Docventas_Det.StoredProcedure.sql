/****** Object:  StoredProcedure [dbo].[spu_Docventas_Mostrar_Docventas_Det]    Script Date: 13/10/2026 21:34:11 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Carga el detalle de documentos, incluyendo la fecha usada por ordenes de compra, sin fijar la base de despliegue.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[spu_Docventas_Mostrar_Docventas_Det]
@idEmpresa     CHAR(2),
@idSucursal    VARCHAR(8),
@idDocumento   CHAR(2),
@idSerie	   CHAR(4),
@iddocventas   CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY

	SELECT idDocumento, idDocVentas, idSerie, idProducto, CAST(glsProducto AS VARCHAR(8000)) AS glsProducto, idMarca, idUM, Factor, Afecto, Cantidad, VVUnit, IGVUnit, PVUnit, TotalVVBruto, 
	TotalPVBruto, PorDcto, DctoVV, DctoPV, TotalVVNeto, TotalIGVNeto, TotalPVNeto, item, GlsMarca, GlsUM, idTipoProducto, idMoneda, idCodFabricante, 
	idEmpresa, idSucursal, estDocImportado, idDocumentoImp, idDocVentasImp, idSerieImp, IdLote,NumLote, FecVencProd, idUsuarioDcto, VVUnitLista, 
	PVUnitLista, CantidadImp, VVUnitNeto, PVUnitNeto, Cantidad2, CodigoRapido, idTallaPeso, CantidadAnt, Simbolo1, Simbolo2, Simbolo3, itemPro, 
	PorcUtilidad, IdCentroCosto, IdSucursalPres, IdDocumentoPres, IdSeriePres, IdDocVentasPres, GlsPlaca, IdUPCliente,IvapUnit,TotalIvapNeto,
	GlsProveedor,CostoS,CostoSInc,CostoD,Margen,VVunitG,
	ISNULL(FechaEmision, GETDATE()) AS FechaEmision,
	ISNULL(Comprado,'0') AS Comprado
	FROM docventasdet d 
	WHERE d.idEmpresa = @idEmpresa 
	AND d.idSucursal = @idSucursal 
	AND d.idDocumento = @idDocumento
	AND d.idDocVentas = @iddocventas
	AND d.idSerie = @idSerie
	ORDER BY ITEM
	;

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
GO
