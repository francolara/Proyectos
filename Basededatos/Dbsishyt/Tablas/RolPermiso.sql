-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define RolPermiso para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.RolPermiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.RolPermiso
    (
        RoleId NVARCHAR(128) NOT NULL,
        CodigoPermiso NVARCHAR(80) NOT NULL,
        CONSTRAINT PK_RolPermiso PRIMARY KEY (RoleId, CodigoPermiso),
        CONSTRAINT FK_RolPermiso_Rol FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id),
        CONSTRAINT FK_RolPermiso_Permiso FOREIGN KEY (CodigoPermiso) REFERENCES dbo.Permiso(Codigo)
    );
END;
GO
