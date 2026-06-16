-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista las empresas activas asociadas al usuario autenticado.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarEmpresasPorUsuario
    @AspNetUserId NVARCHAR(450)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            ue.IdEmpresa,
            e.CodigoEmpresa,
            e.RazonSocial,
            e.NombreComercial,
            e.Ruc,
            ue.EsEmpresaPredeterminada
        FROM dbo.SEG_UsuarioEmpresa AS ue
        INNER JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = ue.IdEmpresa
        WHERE ue.AspNetUserId = @AspNetUserId
          AND ue.Estado = 1
          AND e.Estado = 1
        ORDER BY
            ue.EsEmpresaPredeterminada DESC,
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
