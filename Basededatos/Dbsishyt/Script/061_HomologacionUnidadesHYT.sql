-- Firma: FRANCO LARA | Create date: 11/11/2026 | Completa unidades desde Proveeperu_071126 sin sobrescribir homologaciones existentes.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.UnidadMedida',N'CodigoSunat') IS NULL
    ALTER TABLE dbo.UnidadMedida ADD CodigoSunat NVARCHAR(3) NULL;
GO
SET XACT_ABORT ON;
BEGIN TRANSACTION;

DECLARE @Origen TABLE (Codigo NVARCHAR(30) PRIMARY KEY, Descripcion NVARCHAR(150), CodigoSunat NVARCHAR(3), Alternativo NVARCHAR(30));
INSERT INTO @Origen (Codigo,Descripcion,CodigoSunat,Alternativo) VALUES
(N'CTO',N'Ciento',N'CEN',NULL),
(N'Mil',N'Millar',N'MIL',NULL),
(N'UND',N'Unidad',N'NIU',NULL),
(N'M2',N'M2',N'MTK',NULL),
(N'BAL',N'Balde',N'NIU',NULL),
(N'KL',N'Kilo',N'KGM',NULL),
(N'M3',N'Metro Cubico',N'MTQ',NULL),
(N'PAR',N'Pares',N'PR',NULL),
(N'GLO',N'Paquete',N'PK',NULL),
(N'MT',N'Metros',N'MTR',NULL),
(N'CAJ',N'Caja',N'BX',N'CAJA'),
(N'ROL',N'Rollo',N'RO',NULL),
(N'PLA',N'Plancha',N'NIU',NULL),
(N'PIE',N'Pieza',N'NIU',NULL),
(N'LIS',N'Lister',N'NIU',NULL),
(N'BOL',N'Bolsa',N'BG',N'BLS'),
(N'KG',N'Kilogramo',N'KGM',NULL),
(N'k',N'otro',N'NIU',NULL),
(N'GL',N'GALON',N'GLL',NULL),
(N'FCO',N'FRASCO',N'NIU',NULL),
(N'JGO',N'JUEGO',N'KT',NULL),
(N'P2',N'PIE CUADRADO',N'FOT',NULL),
(N'LT',N'LITRO',N'LTR',NULL),
(N'KIT',N'KIT',N'KT',NULL),
(N'PAQ',N'PAQUETES',N'PK',NULL); -- Misma unidad comercial que Paquete en el origen.

DECLARE @Actor NVARCHAR(128);
SELECT TOP (1) @Actor=usuario.Id
FROM dbo.AspNetUsers AS usuario
WHERE usuario.Activo=1 AND EXISTS (
    SELECT 1 FROM dbo.AspNetUserRoles AS rol
    INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
    WHERE rol.UserId=usuario.Id AND permiso.CodigoPermiso=N'UNIDADES.EDITAR')
ORDER BY usuario.Id;

UPDATE unidad SET CodigoSunat=origen.CodigoSunat,
    FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=COALESCE(@Actor,unidad.UsuarioRegistro)
FROM dbo.UnidadMedida AS unidad
INNER JOIN @Origen AS origen ON unidad.Codigo IN (origen.Codigo,origen.Alternativo)
WHERE NULLIF(TRIM(unidad.CodigoSunat),N'') IS NULL;

IF @Actor IS NOT NULL
    INSERT INTO dbo.UnidadMedida (Codigo,Descripcion,CodigoSunat,Activo,UsuarioRegistro)
    SELECT origen.Codigo,origen.Descripcion,origen.CodigoSunat,1,@Actor
    FROM @Origen AS origen
    WHERE NOT EXISTS (SELECT 1 FROM dbo.UnidadMedida AS unidad WITH (UPDLOCK,HOLDLOCK)
        WHERE unidad.Codigo IN (origen.Codigo,origen.Alternativo));

COMMIT TRANSACTION;
GO
