-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define TarifarioPartida para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioPartida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioPartida
    (
        PartidaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_TarifarioPartida PRIMARY KEY,
        VersionId BIGINT NOT NULL,
        MaterialId BIGINT NOT NULL,
        CodigoMoneda NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        CodigoTipoIGV NVARCHAR(10) NOT NULL,
        PrecioPEN DECIMAL(19,6) NULL,
        PrecioUSD DECIMAL(19,6) NULL,
        Posicion NVARCHAR(60) NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        TipoServicioId INT NULL,
        NumeroNecesidad NVARCHAR(60) NULL,
        CONSTRAINT FK_TarifarioPartida_Version FOREIGN KEY (VersionId) REFERENCES dbo.TarifarioVersion(VersionId),
        CONSTRAINT FK_TarifarioPartida_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_TarifarioPartida_Moneda FOREIGN KEY (CodigoMoneda) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_TarifarioPartida_IGV FOREIGN KEY (CodigoTipoIGV) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT FK_TarifarioPartida_Tipo FOREIGN KEY (TipoServicioId) REFERENCES dbo.TipoServicio(Id),
        CONSTRAINT CK_TarifarioPartida_Precios CHECK ((PrecioPEN IS NULL OR PrecioPEN >= 0) AND (PrecioUSD IS NULL OR PrecioUSD >= 0)),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_TarifarioPartida_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_TarifarioPartida_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_TarifarioPartida_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
