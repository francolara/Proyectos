-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Presupuesto para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Presupuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Presupuesto
    (
        PresupuestoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Presupuesto PRIMARY KEY,
        NumeroInterno NVARCHAR(40) NOT NULL CONSTRAINT UQ_Presupuesto_Numero UNIQUE,
        ClienteId BIGINT NOT NULL,
        TiendaId BIGINT NOT NULL,
        TipoAtencion NVARCHAR(20) NOT NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        FechaSolicitud DATETIME2(3) NULL,
        CONSTRAINT FK_Presupuesto_Tienda FOREIGN KEY (ClienteId, TiendaId) REFERENCES dbo.Tienda(ClienteId, TiendaId),
        CONSTRAINT CK_Presupuesto_Tipo CHECK (TipoAtencion IN (N'INVITACION', N'EMERGENCIA')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Presupuesto_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Presupuesto_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Presupuesto_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
