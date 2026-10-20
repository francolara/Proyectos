-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/10/2026
-- Description:   Lista los niveles configurados para la ayuda principal de productos.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_AyudaProductos_Niveles_Listar
    @IdEmpresa CHAR(2)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            TN.GlsTipoNivel
        FROM dbo.TiposNiveles AS TN
        WHERE TN.IdEmpresa = @IdEmpresa
        ORDER BY TN.Peso ASC;
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
