-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Lista provincias activas por departamento para el mantenimiento de personas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarUbigeoProvincias
    @CodigoDepartamento CHAR(2)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.CodigoProvincia,
            p.CodigoDepartamento,
            p.Nombre
        FROM dbo.UbigeoProvincias AS p
        WHERE p.Activo = 1
          AND p.CodigoDepartamento = @CodigoDepartamento
        ORDER BY
            p.Nombre ASC;

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
