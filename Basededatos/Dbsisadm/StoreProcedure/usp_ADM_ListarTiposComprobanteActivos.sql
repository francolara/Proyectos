-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista tipos de comprobante activos filtrados para compras o ventas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarTiposComprobanteActivos
    @UsoCompras BIT = 0,
    @UsoVentas BIT = 0
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            t.IdTipoComprobante,
            t.CodigoTipoComprobante,
            t.Descripcion,
            t.UsoCompras,
            t.UsoVentas,
            t.Estado
        FROM dbo.ADM_TipoComprobante AS t
        WHERE t.Estado = 1
          AND (@UsoCompras = 0 OR t.UsoCompras = 1)
          AND (@UsoVentas = 0 OR t.UsoVentas = 1)
        ORDER BY
            t.CodigoTipoComprobante ASC;

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
