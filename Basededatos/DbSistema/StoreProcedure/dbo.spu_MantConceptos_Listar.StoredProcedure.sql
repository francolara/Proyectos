-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista conceptos para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantConceptos_Listar
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            C.IdConcepto,
            C.GlsConcepto,
            C.TipoConcepto
        FROM dbo.Conceptos AS C
        WHERE @Busqueda = ''
           OR C.GlsConcepto LIKE '%' + @Busqueda + '%'
           OR C.TipoConcepto LIKE '%' + @Busqueda + '%'
        ORDER BY C.IdConcepto;
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
