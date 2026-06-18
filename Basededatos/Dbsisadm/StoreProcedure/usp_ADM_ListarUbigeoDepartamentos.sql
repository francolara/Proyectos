-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Lista departamentos activos para el mantenimiento de personas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarUbigeoDepartamentos
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            d.CodigoDepartamento,
            d.Nombre
        FROM dbo.UbigeoDepartamentos AS d
        WHERE d.Activo = 1
        ORDER BY
            d.Nombre ASC;

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
