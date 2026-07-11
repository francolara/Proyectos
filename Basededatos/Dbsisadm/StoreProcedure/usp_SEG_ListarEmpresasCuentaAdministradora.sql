-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Lista las empresas activas pertenecientes a una cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarEmpresasCuentaAdministradora
    @IdCuentaAdministradora INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            e.IdEmpresa,
            e.CodigoEmpresa,
            e.RazonSocial,
            e.NombreComercial,
            e.Ruc,
            e.Estado
        FROM dbo.SEG_Empresa AS e
        WHERE e.IdCuentaAdministradora = @IdCuentaAdministradora
          AND e.Estado = 1
        ORDER BY
            e.RazonSocial ASC;

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
