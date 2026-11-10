-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega CCI opcional de veinte digitos a proveedores y trabajadores.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.ProveedorHYT', N'CuentaInterbancaria') IS NULL
    ALTER TABLE dbo.ProveedorHYT ADD CuentaInterbancaria VARCHAR(20) NULL
        CONSTRAINT CK_ProveedorHYT_CCI CHECK (DATALENGTH(CuentaInterbancaria) = 20
            AND CuentaInterbancaria COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^0-9]%');
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'CuentaInterbancaria') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD CuentaInterbancaria VARCHAR(20) NULL
        CONSTRAINT CK_TrabajadorHYT_CCI CHECK (DATALENGTH(CuentaInterbancaria) = 20
            AND CuentaInterbancaria COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^0-9]%');
GO
