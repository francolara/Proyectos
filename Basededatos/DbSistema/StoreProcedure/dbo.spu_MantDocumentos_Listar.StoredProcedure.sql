-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista documentos para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantDocumentos_Listar
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT D.IdDocumento, D.GlsDocumento, D.AbreDocumento
        FROM dbo.Documentos AS D
        WHERE @Busqueda = ''
           OR D.IdDocumento LIKE '%' + @Busqueda + '%'
           OR D.GlsDocumento LIKE '%' + @Busqueda + '%'
           OR D.AbreDocumento LIKE '%' + @Busqueda + '%'
        ORDER BY D.IdDocumento;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
