-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ReferenciaComprobanteExterno para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ReferenciaComprobanteExterno', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ReferenciaComprobanteExterno
    (
        ReferenciaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ReferenciaComprobanteExterno PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL,
        ConfirmacionId BIGINT NOT NULL,
        FechaDocumento DATE NOT NULL,
        CodigoTipoDocumento NVARCHAR(2) NOT NULL,
        Serie NVARCHAR(30) NOT NULL,
        NumeroDocumento NVARCHAR(50) NOT NULL,
        Vigente BIT NOT NULL DEFAULT 1,
        MotivoRetiro NVARCHAR(1000) NULL,
        CONSTRAINT FK_Referencia_Confirmacion FOREIGN KEY (OrdenCompraId, ConfirmacionId) REFERENCES dbo.ConfirmacionFacturacion(OrdenCompraId, ConfirmacionId),
        CONSTRAINT FK_Referencia_Tipo FOREIGN KEY (CodigoTipoDocumento) REFERENCES dbo.TipoDocumento(Codigo),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ReferenciaComprobanteExterno_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ReferenciaComprobanteExterno_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ReferenciaComprobanteExterno_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.ReferenciaComprobanteExterno') AND name = N'UX_Referencia_Vigente')
    CREATE UNIQUE INDEX UX_Referencia_Vigente ON dbo.ReferenciaComprobanteExterno(OrdenCompraId) WHERE Vigente = 1;
GO
