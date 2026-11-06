-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Idempotencia de operaciones documentales, sin duplicar registros por reenvío.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.DocumentoOperacion',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.DocumentoOperacion
    (
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_DocumentoOperacion PRIMARY KEY,
        TipoOperacion NVARCHAR(40) NOT NULL,
        PadreId BIGINT NOT NULL,
        RegistroId BIGINT NOT NULL,
        ArchivoId BIGINT NULL,
        Huella CHAR(64) NOT NULL,
        UsuarioId NVARCHAR(128) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_DocumentoOperacion_Fecha DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_DocumentoOperacion_Usuario FOREIGN KEY (UsuarioId) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_DocumentoOperacion_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO
