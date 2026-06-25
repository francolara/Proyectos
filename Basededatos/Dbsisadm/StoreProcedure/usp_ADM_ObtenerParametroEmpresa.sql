-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Obtiene un parametro especifico de empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ObtenerParametroEmpresa
    @IdEmpresa INT,
    @IdParametroEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            pe.IdParametroEmpresa,
            pe.IdEmpresa,
            pe.TipoParametro,
            pe.CodigoParametro,
            pe.ValorParametro,
            pe.DescripcionParametro,
            pe.FecIni,
            pe.FecFin,
            pe.Activo
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.IdParametroEmpresa = @IdParametroEmpresa;

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
