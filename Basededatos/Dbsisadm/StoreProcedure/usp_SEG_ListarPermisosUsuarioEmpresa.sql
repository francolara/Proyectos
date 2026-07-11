-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Lista permisos efectivos y overrides del usuario para modulos de alcance empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarPermisosUsuarioEmpresa
    @IdUsuarioEmpresa INT
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
            uep.PuedeVer AS PuedeVerOverride,
            uep.PuedeCrear AS PuedeCrearOverride,
            uep.PuedeEditar AS PuedeEditarOverride,
            uep.PuedeEliminar AS PuedeEliminarOverride,
            CAST(COALESCE(uep.PuedeVer, rcp.PuedeVer, 0) AS BIT) AS PuedeVerEfectivo,
            CAST(COALESCE(uep.PuedeCrear, rcp.PuedeCrear, 0) AS BIT) AS PuedeCrearEfectivo,
            CAST(COALESCE(uep.PuedeEditar, rcp.PuedeEditar, 0) AS BIT) AS PuedeEditarEfectivo,
            CAST(COALESCE(uep.PuedeEliminar, rcp.PuedeEliminar, 0) AS BIT) AS PuedeEliminarEfectivo
        FROM dbo.SEG_UsuarioEmpresa AS ue
        INNER JOIN dbo.SEG_UsuarioCuentaAdministradora AS uca
            ON uca.AspNetUserId = ue.AspNetUserId
           AND uca.Estado = 1
        INNER JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = ue.IdEmpresa
           AND e.Estado = 1
           AND e.IdCuentaAdministradora = uca.IdCuentaAdministradora
        INNER JOIN dbo.SEG_ModuloSistema AS ms
            ON ms.AlcanceModulo = 'EMPRESA'
           AND ms.Estado = 1
        LEFT JOIN dbo.SEG_RolCuenta AS rc
            ON rc.CodigoRolCuenta = uca.RolCuenta
           AND rc.Estado = 1
        LEFT JOIN dbo.SEG_RolCuentaPermiso AS rcp
            ON rcp.IdRolCuenta = rc.IdRolCuenta
           AND rcp.IdModuloSistema = ms.IdModuloSistema
           AND rcp.Estado = 1
        LEFT JOIN dbo.SEG_UsuarioEmpresaPermiso AS uep
            ON uep.IdUsuarioEmpresa = ue.IdUsuarioEmpresa
           AND uep.IdModuloSistema = ms.IdModuloSistema
           AND uep.Estado = 1
        WHERE ue.IdUsuarioEmpresa = @IdUsuarioEmpresa
          AND ue.Estado = 1
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
