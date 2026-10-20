-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista grupos de productos para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantGruposProductos_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT GP.IdGrupo, GP.GlsGrupo
        FROM dbo.GruposProducto AS GP
        WHERE GP.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR GP.GlsGrupo LIKE '%' + @Busqueda + '%')
        ORDER BY GP.IdGrupo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
