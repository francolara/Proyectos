-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista marcas para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantMarcas_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            M.IdMarca,
            M.GlsMarca
        FROM dbo.Marcas AS M
        WHERE M.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR M.GlsMarca LIKE '%' + @Busqueda + '%')
        ORDER BY M.IdMarca;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
