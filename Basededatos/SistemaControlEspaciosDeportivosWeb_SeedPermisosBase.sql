-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Carga base de modulos y permisos por rol para administracion de negocio.
-- =============================================

SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.ModulosSistema', N'U') IS NULL OR OBJECT_ID(N'dbo.RolesNegocioPermiso', N'U') IS NULL
BEGIN
    RAISERROR('Primero ejecute el script de estructura AddNegocioPermissions.', 16, 1);
    RETURN;
END;

IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'DASHBOARD')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('DASHBOARD', 'Dashboard', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'USUARIOS_NEGOCIO')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('USUARIOS_NEGOCIO', 'Usuarios del negocio', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'SEDES')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('SEDES', 'Sedes', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'ESPACIOS')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('ESPACIOS', 'Espacios deportivos', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'CLIENTES')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('CLIENTES', 'Clientes', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'RESERVAS')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('RESERVAS', 'Reservas', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'PAGOS')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('PAGOS', 'Pagos', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'TARIFAS')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('TARIFAS', 'Tarifas', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'REPORTES')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('REPORTES', 'Reportes', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = 'COMPROBANTES')
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES ('COMPROBANTES', 'Comprobantes electronicos', 1);

;WITH Roles AS (
    SELECT CAST(1 AS INT) AS RolNegocio UNION ALL -- Administrador
    SELECT 2 UNION ALL                             -- Trabajador
    SELECT 3 UNION ALL                             -- Recepcion
    SELECT 4 UNION ALL                             -- Caja
    SELECT 5                                       -- Supervisor
),
Permisos AS (
    SELECT r.RolNegocio, m.Id AS ModuloSistemaId,
           CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                     WHEN r.RolNegocio = 2 AND m.Codigo IN ('DASHBOARD','RESERVAS','CLIENTES','PAGOS') THEN 1
                     WHEN r.RolNegocio = 3 AND m.Codigo IN ('DASHBOARD','RESERVAS','CLIENTES') THEN 1
                     WHEN r.RolNegocio = 4 AND m.Codigo IN ('DASHBOARD','PAGOS','COMPROBANTES','RESERVAS') THEN 1
                     WHEN r.RolNegocio = 5 AND m.Codigo IN ('DASHBOARD','RESERVAS','REPORTES','ESPACIOS') THEN 1
                     ELSE 0 END AS BIT) AS PuedeVer,
           CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                     WHEN r.RolNegocio = 2 AND m.Codigo IN ('RESERVAS','CLIENTES','PAGOS') THEN 1
                     WHEN r.RolNegocio = 3 AND m.Codigo IN ('RESERVAS','CLIENTES') THEN 1
                     WHEN r.RolNegocio = 4 AND m.Codigo IN ('PAGOS','COMPROBANTES') THEN 1
                     ELSE 0 END AS BIT) AS PuedeCrear,
           CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                     WHEN r.RolNegocio = 2 AND m.Codigo IN ('RESERVAS','CLIENTES','PAGOS') THEN 1
                     WHEN r.RolNegocio = 3 AND m.Codigo IN ('RESERVAS','CLIENTES') THEN 1
                     WHEN r.RolNegocio = 4 AND m.Codigo IN ('PAGOS','COMPROBANTES') THEN 1
                     WHEN r.RolNegocio = 5 AND m.Codigo IN ('RESERVAS','ESPACIOS') THEN 1
                     ELSE 0 END AS BIT) AS PuedeEditar,
           CAST(CASE WHEN r.RolNegocio = 1 THEN 1 ELSE 0 END AS BIT) AS PuedeEliminar
    FROM Roles r
    CROSS JOIN dbo.ModulosSistema m
)
INSERT INTO dbo.RolesNegocioPermiso (RolNegocio, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar)
SELECT p.RolNegocio, p.ModuloSistemaId, p.PuedeVer, p.PuedeCrear, p.PuedeEditar, p.PuedeEliminar
FROM Permisos p
WHERE NOT EXISTS (
    SELECT 1
    FROM dbo.RolesNegocioPermiso rp
    WHERE rp.RolNegocio = p.RolNegocio
      AND rp.ModuloSistemaId = p.ModuloSistemaId
);