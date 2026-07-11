-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Lista las empresas de la cuenta administradora marcando cuales estan asignadas al usuario indicado.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarEmpresasUsuarioCuentaAdministradora
    @IdCuentaAdministradora INT,
    @AspNetUserId NVARCHAR(450)
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
            CAST(CASE WHEN ue.IdUsuarioEmpresa IS NULL THEN 0 ELSE 1 END AS BIT) AS Asignado,
            CAST(COALESCE(ue.EsEmpresaPredeterminada, 0) AS BIT) AS EsEmpresaPredeterminada,
            ue.IdUsuarioEmpresa
        FROM dbo.SEG_Empresa AS e
        LEFT JOIN dbo.SEG_UsuarioEmpresa AS ue
            ON ue.IdEmpresa = e.IdEmpresa
           AND ue.AspNetUserId = @AspNetUserId
           AND ue.Estado = 1
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
