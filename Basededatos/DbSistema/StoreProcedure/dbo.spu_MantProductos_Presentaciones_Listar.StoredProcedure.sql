-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista las presentaciones asociadas a un producto.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProductos_Presentaciones_Listar
    @IdEmpresa CHAR(2),
    @IdProducto CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            P.Item,
            P.IdUM,
            U.GlsUM,
            P.Factor
        FROM dbo.Presentaciones AS P
        INNER JOIN dbo.UnidadMedida AS U
            ON P.IdUM = U.IdUM
        WHERE P.IdEmpresa = @IdEmpresa
          AND P.IdProducto = @IdProducto
        ORDER BY P.Item ASC;
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
