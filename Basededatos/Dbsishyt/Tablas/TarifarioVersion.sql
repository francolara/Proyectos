-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define TarifarioVersion para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioVersion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioVersion
    (
        VersionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_TarifarioVersion PRIMARY KEY,
        TarifarioId BIGINT NOT NULL,
        NumeroVersion INT NOT NULL,
        VigenciaDesde DATE NULL,
        VigenciaHasta DATE NULL,
        CodigoMonedaPredeterminada NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        CodigoTipoIGVPredeterminado NVARCHAR(10) NOT NULL,
        Estado NVARCHAR(20) NOT NULL DEFAULT N'BORRADOR',
        CONSTRAINT UQ_TarifarioVersion_Numero UNIQUE (TarifarioId, NumeroVersion),
        CONSTRAINT FK_TarifarioVersion_Tarifario FOREIGN KEY (TarifarioId) REFERENCES dbo.Tarifario(TarifarioId),
        CONSTRAINT FK_TarifarioVersion_Moneda FOREIGN KEY (CodigoMonedaPredeterminada) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_TarifarioVersion_IGV FOREIGN KEY (CodigoTipoIGVPredeterminado) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT CK_TarifarioVersion_Estado CHECK (Estado IN (N'BORRADOR', N'VIGENTE', N'VENCIDO', N'INACTIVO')),
        CONSTRAINT CK_TarifarioVersion_Fechas CHECK (VigenciaHasta IS NULL OR VigenciaDesde IS NULL OR VigenciaHasta >= VigenciaDesde),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_TarifarioVersion_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_TarifarioVersion_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_TarifarioVersion_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
