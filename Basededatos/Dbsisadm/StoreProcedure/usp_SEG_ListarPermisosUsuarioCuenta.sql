-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Lista permisos efectivos y overrides del usuario para modulos de alcance cuenta.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarPermisosUsuarioCuenta
    @IdUsuarioCuentaAdministradora INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            ms.IdModuloSistema,
            ms.CodigoModulo,
            ms.NombreModulo,
            ms.GrupoMenu,
            uca.RolCuenta,
            rcp.PuedeVer AS PuedeVerRol,
            rcp.PuedeCrear AS PuedeCrearRol,
            rcp.PuedeEditar AS PuedeEditarRol,
            rcp.PuedeEliminar AS PuedeEliminarRol,
            ucp.PuedeVer AS PuedeVerOverride,
            ucp.PuedeCrear AS PuedeCrearOverride,
            ucp.PuedeEditar AS PuedeEditarOverride,
            ucp.PuedeEliminar AS PuedeEliminarOverride,
            CAST(COALESCE(ucp.PuedeVer, rcp.PuedeVer, 0) AS BIT) AS PuedeVerEfectivo,
            CAST(COALESCE(ucp.PuedeCrear, rcp.PuedeCrear, 0) AS BIT) AS PuedeCrearEfectivo,
            CAST(COALESCE(ucp.PuedeEditar, rcp.PuedeEditar, 0) AS BIT) AS PuedeEditarEfectivo,
            CAST(COALESCE(ucp.PuedeEliminar, rcp.PuedeEliminar, 0) AS BIT) AS PuedeEliminarEfectivo
        FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
        INNER JOIN dbo.SEG_ModuloSistema AS ms
            ON ms.AlcanceModulo = 'CUENTA'
           AND ms.Estado = 1
        LEFT JOIN dbo.SEG_RolCuenta AS rc
            ON rc.CodigoRolCuenta = uca.RolCuenta
           AND rc.Estado = 1
        LEFT JOIN dbo.SEG_RolCuentaPermiso AS rcp
            ON rcp.IdRolCuenta = rc.IdRolCuenta
           AND rcp.IdModuloSistema = ms.IdModuloSistema
           AND rcp.Estado = 1
        LEFT JOIN dbo.SEG_UsuarioCuentaPermiso AS ucp
            ON ucp.IdUsuarioCuentaAdministradora = uca.IdUsuarioCuentaAdministradora
           AND ucp.IdModuloSistema = ms.IdModuloSistema
           AND ucp.Estado = 1
        WHERE uca.IdUsuarioCuentaAdministradora = @IdUsuarioCuentaAdministradora
          AND uca.Estado = 1
        ORDER BY
            ms.OrdenMenu ASC,
            ms.NombreModulo ASC;

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
