-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista zonas para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantZonas_Listar
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT Z.IdZona, Z.GlsZona
        FROM dbo.Zonas AS Z
        WHERE @Busqueda = '' OR Z.GlsZona LIKE '%' + @Busqueda + '%'
        ORDER BY Z.IdZona;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
