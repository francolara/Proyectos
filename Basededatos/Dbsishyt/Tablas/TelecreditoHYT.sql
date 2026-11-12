-- Firma: FRANCO LARA | Create date: 10/11/2026 | Cuenta de cargo BCP de la empresa.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TelecreditoHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TelecreditoHYT
    (
        Id TINYINT NOT NULL CONSTRAINT PK_TelecreditoHYT PRIMARY KEY CHECK (Id=1),
        TipoCuenta CHAR(1) NOT NULL CHECK (TipoCuenta IN ('C','M')),
        NumeroCuenta VARCHAR(13) NOT NULL CHECK (LEN(NumeroCuenta)=13
            AND NumeroCuenta COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^0-9]%' AND SUBSTRING(NumeroCuenta,11,1)='0'),
        Moneda CHAR(3) NOT NULL CHECK (Moneda='PEN'),
        VersionFila ROWVERSION NOT NULL
    );
END;
GO
