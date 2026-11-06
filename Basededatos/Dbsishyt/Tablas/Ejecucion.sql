-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Ejecucion para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Ejecucion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Ejecucion
    (
        EjecucionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Ejecucion PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL,
        Responsable NVARCHAR(150) NULL,
        FechaProgramada DATETIME2(3) NULL,
        FechaInicioReal DATETIME2(3) NULL,
        FechaTerminoReal DATETIME2(3) NULL,
        Descripcion NVARCHAR(2000) NULL,
        Estado NVARCHAR(30) NOT NULL DEFAULT N'PENDIENTE',
        CONSTRAINT FK_Ejecucion_Presupuesto FOREIGN KEY (PresupuestoId) REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT CK_Ejecucion_Estado CHECK (Estado IN (N'PENDIENTE', N'PROGRAMADA', N'EN_EJECUCION', N'TERMINADA', N'CANCELADA')),
        CONSTRAINT CK_Ejecucion_Fechas CHECK (FechaTerminoReal IS NULL OR FechaInicioReal IS NULL OR FechaTerminoReal >= FechaInicioReal),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Ejecucion_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Ejecucion_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Ejecucion_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
