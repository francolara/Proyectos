-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Cliente para la base inicial de HYT.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Código SAP opcional por cliente, independiente del código interno.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Cliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Cliente
    (
        ClienteId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Cliente PRIMARY KEY,
        Codigo NVARCHAR(40) NOT NULL CONSTRAINT UQ_Cliente_Codigo UNIQUE,
        CodigoSAP NVARCHAR(60) NULL,
        TipoEntidad NVARCHAR(20) NOT NULL,
        TipoIdentificacion NVARCHAR(20) NULL,
        NumeroIdentificacion NVARCHAR(30) NULL,
        RazonSocial NVARCHAR(200) NULL,
        Nombres NVARCHAR(150) NULL,
        Apellidos NVARCHAR(150) NULL,
        DireccionFiscal NVARCHAR(300) NULL,
        UrlCargaDocumentos NVARCHAR(2048) NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT CK_Cliente_Entidad CHECK (TipoEntidad IN (N'JURIDICA', N'NATURAL')),
        CONSTRAINT CK_Cliente_Nombre CHECK ((TipoEntidad = N'JURIDICA' AND NULLIF(LTRIM(RTRIM(RazonSocial)), N'') IS NOT NULL) OR
            (TipoEntidad = N'NATURAL' AND NULLIF(LTRIM(RTRIM(Nombres)), N'') IS NOT NULL AND NULLIF(LTRIM(RTRIM(Apellidos)), N'') IS NOT NULL)),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Cliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Cliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Cliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
