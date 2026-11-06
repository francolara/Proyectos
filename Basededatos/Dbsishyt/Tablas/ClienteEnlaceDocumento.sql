-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Enlaces de documentos con descripción, orden y responsables por cliente.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ClienteEnlaceDocumento',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ClienteEnlaceDocumento
    (
        Id BIGINT IDENTITY NOT NULL CONSTRAINT PK_ClienteEnlaceDocumento PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        Descripcion NVARCHAR(200) NOT NULL,
        Url NVARCHAR(2048) NOT NULL,
        Orden INT NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ClienteEnlaceDocumento_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        CONSTRAINT FK_ClienteEnlaceDocumento_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT FK_ClienteEnlaceDocumento_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ClienteEnlaceDocumento_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
    CREATE INDEX IX_ClienteEnlaceDocumento_Cliente ON dbo.ClienteEnlaceDocumento(ClienteId,Orden,Id);
END;
GO
