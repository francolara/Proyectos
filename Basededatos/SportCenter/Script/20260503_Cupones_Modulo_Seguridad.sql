USE [DbSportCenter]
GO
-- Firma: FRANCO LARA
-- Create date: 03/05/2026
-- Descripcion: registra modulo CUPONES y permisos base por rol de negocio.
IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'CUPONES')
BEGIN
    INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo)
    VALUES (N'CUPONES', N'Cupones', 1);
END
GO
;WITH Roles AS
(
    SELECT CAST(1 AS INT) AS RolNegocio UNION ALL
    SELECT 2 UNION ALL
    SELECT 3 UNION ALL
    SELECT 4 UNION ALL
    SELECT 5
)
INSERT INTO dbo.RolesNegocioPermiso (RolNegocio, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar)
SELECT
    r.RolNegocio,
    m.Id,
    CAST(CASE WHEN r.RolNegocio IN (1,2,3) THEN 1 ELSE 0 END AS BIT),
    CAST(CASE WHEN r.RolNegocio IN (1,2) THEN 1 ELSE 0 END AS BIT),
    CAST(CASE WHEN r.RolNegocio IN (1,2) THEN 1 ELSE 0 END AS BIT),
    CAST(CASE WHEN r.RolNegocio = 1 THEN 1 ELSE 0 END AS BIT)
FROM Roles r
INNER JOIN dbo.ModulosSistema m ON m.Codigo = N'CUPONES'
WHERE NOT EXISTS (
    SELECT 1
    FROM dbo.RolesNegocioPermiso rp
    WHERE rp.RolNegocio = r.RolNegocio
      AND rp.ModuloSistemaId = m.Id
);
GO
