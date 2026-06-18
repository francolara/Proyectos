-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista las monedas activas del sistema para operaciones administrativas y contables.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarMonedasActivas
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            m.IdMoneda,
            m.CodigoMoneda,
            m.NombreMoneda,
            m.SimboloMoneda,
            m.EsMonedaBase,
            m.Estado
        FROM dbo.ADM_Moneda AS m
        WHERE m.Estado = 1
        ORDER BY
            m.EsMonedaBase DESC,
            m.CodigoMoneda ASC;

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
