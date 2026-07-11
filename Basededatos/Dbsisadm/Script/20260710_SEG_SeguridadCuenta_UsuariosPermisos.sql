-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Ejecuta la semilla base de modulos, roles y permisos para seguridad por opcion, y regulariza usuarios fundadores grabados con rol ADMINISTRADOR.
-- =============================================

IF OBJECT_ID(N'dbo.usp_SEG_SeedSeguridadCuentaPermisosBase', N'P') IS NOT NULL
BEGIN
    EXEC dbo.usp_SEG_SeedSeguridadCuentaPermisosBase
        @UsuarioRegistro = N'sistema';
END;

UPDATE uca
SET uca.RolCuenta = N'ADMINISTRADORCUENTA',
    uca.UsuarioRegistro = N'sistema'
FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
WHERE uca.RolCuenta = N'ADMINISTRADOR';
