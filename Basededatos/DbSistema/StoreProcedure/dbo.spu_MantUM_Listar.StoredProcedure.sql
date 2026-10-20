-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista unidades de medida para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUM_Listar
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            UM.IdUM,
            UM.GlsUM,
            UM.AbreUM
        FROM dbo.UnidadMedida AS UM
        WHERE @Busqueda = ''
           OR UM.GlsUM LIKE '%' + @Busqueda + '%'
           OR UM.AbreUM LIKE '%' + @Busqueda + '%'
        ORDER BY UM.IdUM;
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
