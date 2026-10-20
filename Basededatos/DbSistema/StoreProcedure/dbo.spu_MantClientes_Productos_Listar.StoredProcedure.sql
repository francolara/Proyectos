-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista los codigos de productos asociados al cliente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_Productos_Listar
    @IdEmpresa CHAR(2),
    @IdCliente CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PC.IdEmpresa,
            PC.IdProducto,
            P.GlsProducto,
            PC.Codigo,
            PC.CodigoBarra
        FROM dbo.ProductosClientes AS PC
        INNER JOIN dbo.Productos AS P
            ON PC.IdEmpresa = P.IdEmpresa
           AND PC.IdProducto = P.IdProducto
        WHERE PC.IdEmpresa = @IdEmpresa
          AND PC.IdCliente = @IdCliente
        ORDER BY PC.IdProducto;
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
