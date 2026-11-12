-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registra permisos de negociación de factoring para Administración.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
INSERT INTO dbo.Permiso(Codigo,Descripcion)
SELECT nuevo.Codigo,nuevo.Descripcion
FROM (VALUES (N'FACTORING.CONSULTAR',N'Consultar negociaciones de factoring'),
             (N'FACTORING.EDITAR',N'Registrar negociaciones de factoring')) AS nuevo(Codigo,Descripcion)
WHERE NOT EXISTS(SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo=nuevo.Codigo);

INSERT INTO dbo.RolPermiso(RoleId,CodigoPermiso)
SELECT rol.Id,permiso.Codigo FROM dbo.AspNetRoles AS rol CROSS JOIN dbo.Permiso AS permiso
WHERE rol.Id IN(N'HYT-ADMIN',N'HYT-ADMINISTRACION') AND permiso.Codigo IN(N'FACTORING.CONSULTAR',N'FACTORING.EDITAR')
    AND NOT EXISTS(SELECT 1 FROM dbo.RolPermiso AS asignado WHERE asignado.RoleId=rol.Id AND asignado.CodigoPermiso=permiso.Codigo);
GO
