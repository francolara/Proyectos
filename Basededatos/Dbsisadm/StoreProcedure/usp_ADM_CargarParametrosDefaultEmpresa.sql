-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia parametros maestros internos hacia una empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_CargarParametrosDefaultEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        INSERT INTO dbo.ADM_ParametroEmpresa
        (
            IdEmpresa,
            TipoParametro,
            CodigoParametro,
            ValorParametro,
            DescripcionParametro,
            FecIni,
            FecFin,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            pm.TipoParametro,
            pm.CodigoParametro,
            pm.ValorParametro,
            pm.DescripcionParametro,
            pm.FecIni,
            pm.FecFin,
            pm.Activo,
            @UsuarioRegistro
        FROM dbo.ADM_ParametroMaestro AS pm
        WHERE pm.Activo = 1
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.ADM_ParametroEmpresa AS pe
              WHERE pe.IdEmpresa = @IdEmpresa
                AND pe.TipoParametro = pm.TipoParametro
                AND pe.CodigoParametro = pm.CodigoParametro
          );

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
