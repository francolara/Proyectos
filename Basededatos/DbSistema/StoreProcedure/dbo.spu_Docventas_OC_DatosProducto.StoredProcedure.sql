-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Obtiene los datos base de un producto usado en ordenes de compra.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_OC_DatosProducto
    @IdEmpresa CHAR(2),
    @IdProducto VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            P.IdFabricante,
            P.IdMarca,
            M.GlsMarca,
            P.AfectoIGV,
            P.IdTipoProducto
        FROM dbo.Productos AS P
        LEFT JOIN dbo.Marcas AS M
            ON M.IdEmpresa = P.IdEmpresa
           AND M.IdMarca = P.IdMarca
        WHERE P.IdEmpresa = @IdEmpresa
          AND P.IdProducto = @IdProducto;
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
