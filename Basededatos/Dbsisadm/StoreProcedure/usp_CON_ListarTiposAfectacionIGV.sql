-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Lista tipos de afectacion IGV activos para registros de provision.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarTiposAfectacionIGV
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            t.IdTipoAfectacionIGV,
            t.CodigoSunat,
            t.NombreAfectacion,
            t.Estado
        FROM dbo.CON_TipoAfectacionIGV AS t
        WHERE t.Estado = 1
        ORDER BY
            t.CodigoSunat ASC;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
