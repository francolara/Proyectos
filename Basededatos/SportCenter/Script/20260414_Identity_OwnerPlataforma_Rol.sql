USE [DbSportCenter];
GO

-- Firma: Codex - 14/04/2026 | Crea rol global OwnerPlataforma y permite asignarlo a un usuario Identity por correo.

IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles WHERE [Name] = N'OwnerPlataforma')
BEGIN
    INSERT INTO dbo.AspNetRoles (Id, [Name], NormalizedName, ConcurrencyStamp)
    VALUES (CONVERT(NVARCHAR(450), NEWID()), N'OwnerPlataforma', N'OWNERPLATAFORMA', CONVERT(NVARCHAR(36), NEWID()));
END;
GO

DECLARE @Correo NVARCHAR(256) = N'llara@efc.com.pe'; -- Cambiar por el correo del dueno de la plataforma.

DECLARE @UserId NVARCHAR(450) = (
    SELECT TOP (1) u.Id
    FROM dbo.AspNetUsers u
    WHERE u.NormalizedEmail = UPPER(LTRIM(RTRIM(@Correo)))
       OR u.Email = LTRIM(RTRIM(@Correo))
);

DECLARE @RoleId NVARCHAR(450) = (
    SELECT TOP (1) r.Id
    FROM dbo.AspNetRoles r
    WHERE r.NormalizedName = N'OWNERPLATAFORMA'
);

IF @UserId IS NULL
BEGIN
    RAISERROR(N'No existe usuario para el correo indicado.', 16, 1);
    RETURN;
END;

IF @RoleId IS NULL
BEGIN
    RAISERROR(N'No existe el rol OwnerPlataforma.', 16, 1);
    RETURN;
END;

IF NOT EXISTS (
    SELECT 1
    FROM dbo.AspNetUserRoles ur
    WHERE ur.UserId = @UserId
      AND ur.RoleId = @RoleId
)
BEGIN
    INSERT INTO dbo.AspNetUserRoles (UserId, RoleId)
    VALUES (@UserId, @RoleId);
END;
GO
