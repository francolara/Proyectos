-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista los almacenes asociados a un producto.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProductos_Almacenes_Listar
    @IdEmpresa CHAR(2),
    @IdProducto CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PA.IdUbicacion,
            PA.Item,
            PA.IdSucursal,
            S.GlsPersona AS GlsSucursal,
            PA.IdAlmacen,
            A.GlsAlmacen
        FROM dbo.ProductosAlmacen AS PA
        INNER JOIN dbo.Almacenes AS A
            ON PA.IdEmpresa = A.IdEmpresa
           AND PA.IdAlmacen = A.IdAlmacen
        INNER JOIN dbo.Personas AS S
            ON PA.IdSucursal = S.IdPersona
        WHERE PA.IdEmpresa = @IdEmpresa
          AND PA.IdProducto = @IdProducto
        ORDER BY PA.Item;
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
