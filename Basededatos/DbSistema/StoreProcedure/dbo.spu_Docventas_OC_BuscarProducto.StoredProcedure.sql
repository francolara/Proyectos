-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Resuelve un producto de OC por codigo, fabricante o codigo rapido.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_OC_BuscarProducto
    @IdEmpresa CHAR(2),
    @Codigo VARCHAR(40)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT TOP (1)
            P.IdProducto,
            P.CodigoRapido,
            P.GlsProducto,
            P.IdUMVenta,
            UM.AbreUM
        FROM dbo.Productos AS P
        INNER JOIN dbo.UnidadMedida AS UM
            ON UM.IdUM = P.IdUMVenta
        WHERE P.IdEmpresa = @IdEmpresa
          AND (
                P.IdProducto = @Codigo
                OR P.IdFabricante = @Codigo
                OR P.CodigoRapido = @Codigo
              )
        ORDER BY
            CASE WHEN P.IdProducto = @Codigo THEN 0
                 WHEN P.CodigoRapido = @Codigo THEN 1
                 ELSE 2
            END,
            P.IdProducto;
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
