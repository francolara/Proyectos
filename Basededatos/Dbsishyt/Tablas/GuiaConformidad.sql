-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define GuiaConformidad para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.GuiaConformidad', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GuiaConformidad
    (
        GuiaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_GuiaConformidad PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL CONSTRAINT UQ_Guia_Presupuesto UNIQUE,
        FechaActa DATE NOT NULL,
        NumeroActa NVARCHAR(100) NOT NULL,
        Observacion NVARCHAR(2000) NULL,
        Estado NVARCHAR(20) NOT NULL DEFAULT N'REGISTRADA',
        FechaAceptacion DATETIME2(3) NULL,
        RepresentanteCliente NVARCHAR(150) NULL,
        CONSTRAINT FK_Guia_Presupuesto FOREIGN KEY (PresupuestoId) REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT CK_Guia_Estado CHECK (Estado IN (N'REGISTRADA', N'ENVIADA', N'OBSERVADA', N'ACEPTADA', N'ANULADA')),
        CONSTRAINT CK_Guia_Aceptacion CHECK (Estado <> N'ACEPTADA' OR (FechaAceptacion IS NOT NULL AND RepresentanteCliente IS NOT NULL)),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_GuiaConformidad_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_GuiaConformidad_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_GuiaConformidad_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
