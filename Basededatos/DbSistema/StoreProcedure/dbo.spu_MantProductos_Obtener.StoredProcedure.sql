-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene la cabecera de un producto para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProductos_Obtener
    @IdEmpresa CHAR(2),
    @IdProducto CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            P.IdProducto,
            P.GlsProducto,
            P.IdNivel,
            P.IdUMCompra,
            P.IdUMVenta,
            P.IdMarca,
            P.AfectoIGV,
            P.IdMoneda,
            P.IdTipoProducto,
            P.IdFabricante,
            P.CodigoRapido,
            P.IdGrupo,
            P.IdTallaPeso,
            P.IndImagen,
            P.IndInsertaPrecioLista,
            P.AfectoDctoEspecial,
            P.CtaContable,
            P.EstProducto,
            P.CtaContable2,
            P.CtaContable_Relacionada,
            P.TiempoRepInv,
            P.IdConceptoDetraccion,
            ISNULL(P.IndRptGL, '0') AS IndRptGL,
            P.StockMinimo,
            CAST('0' + ISNULL(P.IndAfectoIVAP, 0) AS TINYINT) AS IndAfectoIVAP,
            P.IdConceptoCosteo
        FROM dbo.Productos AS P
        WHERE P.IdEmpresa = @IdEmpresa
          AND P.IdProducto = @IdProducto;
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
