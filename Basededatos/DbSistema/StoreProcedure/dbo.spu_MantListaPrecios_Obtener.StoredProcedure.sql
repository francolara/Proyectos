-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene la cabecera de una lista de precios.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantListaPrecios_Obtener
    @IdEmpresa CHAR(2),
    @IdLista CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            LP.IdLista,
            LP.GlsLista,
            LP.FecVcto,
            LP.EstLista,
            LP.IdMoneda
        FROM dbo.ListaPrecios AS LP
        WHERE LP.IdEmpresa = @IdEmpresa
          AND LP.IdLista = @IdLista;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
