-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/10/2026
-- Description:   Devuelve la configuracion de busqueda del vendedor al iniciar sesion.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Usuario_ConfiguracionBusqueda
    @IdUsuario CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            ISNULL(MAX(CASE WHEN CONVERT(VARCHAR(10), V.IndVisualizaClientes) = '1' THEN 1 ELSE 0 END), 0) AS IndVisualizaClientes
        FROM dbo.Vendedores AS V
        WHERE V.IdVendedor = @IdUsuario;
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
