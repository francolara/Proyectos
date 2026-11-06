-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ConfirmacionFacturacion para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ConfirmacionFacturacion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConfirmacionFacturacion
    (
        ConfirmacionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ConfirmacionFacturacion PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL,
        FechaConfirmacion DATE NOT NULL,
        NumeroConfirmacion NVARCHAR(100) NOT NULL,
        Observacion NVARCHAR(2000) NULL,
        RepresentanteCliente NVARCHAR(150) NULL,
        Vigente BIT NOT NULL DEFAULT 1,
        CONSTRAINT UQ_Confirmacion_Vinculo UNIQUE (OrdenCompraId, ConfirmacionId),
        CONSTRAINT FK_Confirmacion_OC FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ConfirmacionFacturacion_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ConfirmacionFacturacion_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ConfirmacionFacturacion_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.ConfirmacionFacturacion') AND name = N'UX_Confirmacion_Vigente')
    CREATE UNIQUE INDEX UX_Confirmacion_Vigente ON dbo.ConfirmacionFacturacion(OrdenCompraId) WHERE Vigente = 1;
GO
