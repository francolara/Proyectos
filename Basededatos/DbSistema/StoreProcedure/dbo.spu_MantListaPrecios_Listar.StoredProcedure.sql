-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista listas de precios para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantListaPrecios_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            LP.IdLista,
            LP.GlsLista,
            LP.FecVcto,
            LP.EstLista
        FROM dbo.ListaPrecios AS LP
        WHERE LP.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR LP.GlsLista LIKE '%' + @Busqueda + '%')
        ORDER BY LP.IdLista;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
