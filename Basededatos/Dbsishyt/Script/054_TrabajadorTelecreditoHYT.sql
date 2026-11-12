-- Firma: FRANCO LARA | Create date: 10/11/2026 | Datos opcionales de identidad y abono BCP. Admite cuenta interbancaria sin número local.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.TrabajadorHYT',N'TipoDocumentoBcp') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD TipoDocumentoBcp CHAR(1) NULL
        CONSTRAINT CK_TrabajadorHYT_TipoDocumentoBcp CHECK (TipoDocumentoBcp IN ('1','3','4'));
IF COL_LENGTH(N'dbo.TrabajadorHYT',N'DocumentoBcp') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD DocumentoBcp VARCHAR(12) NULL;
IF COL_LENGTH(N'dbo.TrabajadorHYT',N'TipoCuentaBcp') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD TipoCuentaBcp CHAR(1) NULL
        CONSTRAINT CK_TrabajadorHYT_TipoCuentaBcp CHECK (TipoCuentaBcp IN ('A','C','M','B'));
IF COL_LENGTH(N'dbo.TrabajadorHYT',N'MonedaCuentaBcp') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD MonedaCuentaBcp CHAR(3) NULL
        CONSTRAINT CK_TrabajadorHYT_MonedaCuentaBcp CHECK (MonedaCuentaBcp IN ('PEN','USD'));
GO
IF OBJECT_ID(N'dbo.CK_TrabajadorHYT_CuentaOCci',N'C') IS NULL
BEGIN
    IF OBJECT_ID(N'dbo.CK_TrabajadorHYT_Cuenta',N'C') IS NOT NULL
        ALTER TABLE dbo.TrabajadorHYT DROP CONSTRAINT CK_TrabajadorHYT_Cuenta;
    ALTER TABLE dbo.TrabajadorHYT ADD CONSTRAINT CK_TrabajadorHYT_CuentaOCci
        CHECK ((Banco IS NULL AND NumeroCuenta IS NULL AND CuentaInterbancaria IS NULL)
            OR (Banco IS NOT NULL AND LEN(Banco)>0
                AND (ISNULL(LEN(NumeroCuenta),0)>0 OR CuentaInterbancaria IS NOT NULL)));
END;
GO
