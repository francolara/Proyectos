-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ImpresionPresupuesto para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ImpresionPresupuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ImpresionPresupuesto
    (
        ImpresionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ImpresionPresupuesto PRIMARY KEY,
        RevisionId BIGINT NOT NULL,
        Formato NVARCHAR(10) NOT NULL,
        VersionFormato NVARCHAR(20) NOT NULL,
        CONSTRAINT FK_Impresion_Revision FOREIGN KEY (RevisionId) REFERENCES dbo.PresupuestoRevision(RevisionId),
        CONSTRAINT CK_Impresion_Formato CHECK (Formato IN (N'ANTERIOR', N'NUEVO')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ImpresionPresupuesto_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ImpresionPresupuesto_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ImpresionPresupuesto_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
