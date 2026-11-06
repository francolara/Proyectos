-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define SeguimientoCliente para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.SeguimientoCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SeguimientoCliente
    (
        SeguimientoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_SeguimientoCliente PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL,
        FechaEvento DATETIME2(3) NOT NULL,
        Observacion NVARCHAR(2000) NOT NULL,
        CONSTRAINT FK_Seguimiento_OC FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_SeguimientoCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_SeguimientoCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_SeguimientoCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
