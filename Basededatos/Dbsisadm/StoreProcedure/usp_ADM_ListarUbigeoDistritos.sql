-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Lista distritos activos por provincia para el mantenimiento de personas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarUbigeoDistritos
    @CodigoProvincia CHAR(4)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            d.CodigoUbigeo,
            d.CodigoDepartamento,
            d.CodigoProvincia,
            d.Nombre,
            d.Zona
        FROM dbo.UbigeoDistritos AS d
        WHERE d.Activo = 1
          AND d.CodigoProvincia = @CodigoProvincia
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
