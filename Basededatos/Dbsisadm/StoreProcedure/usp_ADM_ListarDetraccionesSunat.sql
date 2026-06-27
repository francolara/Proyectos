-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Lista los codigos SUNAT de detraccion activos para ayudas y operaciones de compras.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarDetraccionesSunat
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            d.IdDetraccionSunat,
            d.CodigoSunat,
            d.Descripcion,
            d.Porcentaje,
            d.Estado
        FROM dbo.ADM_DetraccionSunat AS d
        WHERE d.Estado = 1
        ORDER BY
            d.CodigoSunat ASC;

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
