-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define ProveedorCliente para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ProveedorCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ProveedorCliente
    (
        Id BIGINT IDENTITY NOT NULL CONSTRAINT PK_ProveedorCliente PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        CodigoSAPProveedor NVARCHAR(60) NULL,
        ContratoMarco NVARCHAR(60) NULL,
        VigenciaDesde DATE NULL,
        VigenciaHasta DATE NULL,
        CONSTRAINT FK_ProveedorCliente_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT CK_ProveedorCliente_Fechas CHECK (VigenciaHasta IS NULL OR VigenciaDesde IS NULL OR VigenciaHasta >= VigenciaDesde),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ProveedorCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ProveedorCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ProveedorCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
