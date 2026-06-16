-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Tipos de cambio diarios por moneda.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_TipoCambio', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_TipoCambio
    (
        IdTipoCambio INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_TipoCambio PRIMARY KEY,
        IdMoneda INT NOT NULL,
        Fecha DATE NOT NULL,
        TipoCambioCompra DECIMAL(18,6) NOT NULL,
        TipoCambioVenta DECIMAL(18,6) NOT NULL,
        TipoCambioFacturacion DECIMAL(18,6) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_ADM_TipoCambio_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_TipoCambio_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_TipoCambio
        ADD CONSTRAINT FK_ADM_TipoCambio_ADM_Moneda
        FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.ADM_TipoCambio
        ADD CONSTRAINT UQ_ADM_TipoCambio_MonedaFecha UNIQUE (IdMoneda, Fecha);
END;
