-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Lista tipos de documento de identidad SUNAT activos para formularios administrativos.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarTiposDocumentoIdentidadSunat
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            t.CodigoSunat,
            t.CodigoInterno,
            t.Nombre,
            t.Orden
        FROM dbo.TiposDocumentoIdentidadSunat AS t
        WHERE t.Activo = 1
        ORDER BY
            t.Orden ASC,
            t.Nombre ASC;

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
