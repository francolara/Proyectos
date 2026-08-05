-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Obtiene la ultima huella generada para decidir si el plan PLE debe llevar informacion.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_PlanContable_ControlObtener
    @IdEmpresa INT,
    @Anio SMALLINT,
    @CodigoFormato VARCHAR(10)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            pc.IdEmpresa,
            pc.Anio,
            pc.CodigoFormato,
            pc.HuellaPlanContable,
            pc.FechaUltimaGeneracion
        FROM dbo.CON_PLE_PlanContableControl AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.Anio = @Anio
          AND pc.CodigoFormato = @CodigoFormato;

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
