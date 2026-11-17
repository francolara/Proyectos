-- Firma: FRANCO LARA | Create date: 17/11/2026 | Catálogo de firmas nombradas con una sola predeterminada.
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.FirmaPresupuestoHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FirmaPresupuestoHYT
    (
        FirmaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_FirmaPresupuestoHYT PRIMARY KEY,
        Nombre NVARCHAR(150) NOT NULL CONSTRAINT UQ_FirmaPresupuestoHYT_Nombre UNIQUE,
        ArchivoJson NVARCHAR(2000) NOT NULL CHECK(ISJSON(ArchivoJson)=1),
        Predeterminada BIT NOT NULL DEFAULT 0,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL REFERENCES dbo.AspNetUsers(Id)
    );
    CREATE UNIQUE INDEX UX_FirmaPresupuestoHYT_Predeterminada ON dbo.FirmaPresupuestoHYT(Predeterminada) WHERE Predeterminada=1;
END;
GO
