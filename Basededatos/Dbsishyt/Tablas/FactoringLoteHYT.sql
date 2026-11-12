-- Firma: FRANCO LARA | Create date: 10/11/2026 | Cabecera de negociación de factoring.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.FactoringLoteHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FactoringLoteHYT
    (
        LoteId BIGINT IDENTITY NOT NULL CONSTRAINT PK_FactoringLoteHYT PRIMARY KEY,
        Financiera NVARCHAR(150) NOT NULL,
        FechaNegociacion DATE NOT NULL,
        FechaDesembolso DATE NULL,
        Porcentaje DECIMAL(7,4) NOT NULL CHECK (Porcentaje > 0 AND Porcentaje <= 100),
        Moneda NVARCHAR(3) NOT NULL REFERENCES dbo.Moneda(Codigo),
        Referencia NVARCHAR(100) NULL,
        Observaciones NVARCHAR(500) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_FactoringLoteHYT_Operacion UNIQUE,
        Huella VARBINARY(32) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_FactoringLoteHYT_Fechas CHECK (FechaDesembolso IS NULL OR FechaDesembolso >= FechaNegociacion)
    );
END;
GO
