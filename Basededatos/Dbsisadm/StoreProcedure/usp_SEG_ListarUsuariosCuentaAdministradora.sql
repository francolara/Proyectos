-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Lista usuarios activos de la cuenta administradora con resumen de empresas asignadas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarUsuariosCuentaAdministradora
    @IdCuentaAdministradora INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            uca.IdUsuarioCuentaAdministradora,
            uca.AspNetUserId,
            ISNULL(au.Email, au.UserName) AS CorreoUsuario,
            up.NombreCompleto,
            up.Telefono,
            uca.RolCuenta,
            uca.EsCuentaPredeterminada,
            uca.Estado,
            COUNT(e.IdEmpresa) AS CantidadEmpresasAsignadas,
            STRING_AGG(CONCAT(e.CodigoEmpresa, N' - ', e.RazonSocial), N' | ') WITHIN GROUP (ORDER BY e.RazonSocial) AS EmpresasAsignadas
        FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
        INNER JOIN dbo.AspNetUsers AS au
            ON au.Id = uca.AspNetUserId
        LEFT JOIN dbo.SEG_UsuarioPerfil AS up
            ON up.AspNetUserId = uca.AspNetUserId
        LEFT JOIN dbo.SEG_UsuarioEmpresa AS ue
            ON ue.AspNetUserId = uca.AspNetUserId
           AND ue.Estado = 1
        LEFT JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = ue.IdEmpresa
           AND e.Estado = 1
           AND e.IdCuentaAdministradora = @IdCuentaAdministradora
        WHERE uca.IdCuentaAdministradora = @IdCuentaAdministradora
          AND uca.Estado = 1
        GROUP BY
            uca.IdUsuarioCuentaAdministradora,
            uca.AspNetUserId,
            au.Email,
            au.UserName,
            up.NombreCompleto,
            up.Telefono,
            uca.RolCuenta,
            uca.EsCuentaPredeterminada,
            uca.Estado
        ORDER BY
            uca.RolCuenta ASC,
            ISNULL(up.NombreCompleto, ISNULL(au.Email, au.UserName)) ASC;

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
