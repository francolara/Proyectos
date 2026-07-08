-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Crea el mantenimiento de tipos de cambio por cuenta administradora.
-- =============================================
-- Firma: FRANCO LARA - 29/06/2026 | Crea la tabla CON_TipoCambio por cuenta administradora con unicidad por fecha y moneda.

IF OBJECT_ID(N'dbo.CON_TipoCambio', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoCambio
    (
        IdTipoCambio INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoCambio PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        Fecha DATE NOT NULL,
        IdMoneda VARCHAR(3) NOT NULL,
        Compra DECIMAL(18,4) NOT NULL,
        Venta DECIMAL(18,4) NOT NULL,
        CompraSBS DECIMAL(18,4) NOT NULL,
        VentaSBS DECIMAL(18,4) NOT NULL,
        Fuente VARCHAR(50) NOT NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_TipoCambio_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_TipoCambio_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_TipoCambio
        ADD CONSTRAINT FK_CON_TipoCambio_SEG_CuentaAdministradora
            FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);

    ALTER TABLE dbo.CON_TipoCambio
        ADD CONSTRAINT UQ_CON_TipoCambio_Cuenta_Fecha_Moneda
            UNIQUE (IdCuentaAdministradora, Fecha, IdMoneda);

    ALTER TABLE dbo.CON_TipoCambio
        ADD CONSTRAINT CK_CON_TipoCambio_Montos
            CHECK (Compra > 0 AND Venta > 0 AND CompraSBS > 0 AND VentaSBS > 0);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_CON_TipoCambio_SEG_CuentaAdministradora'
)
BEGIN
    ALTER TABLE dbo.CON_TipoCambio
        ADD CONSTRAINT FK_CON_TipoCambio_SEG_CuentaAdministradora
            FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.key_constraints
    WHERE name = N'UQ_CON_TipoCambio_Cuenta_Fecha_Moneda'
)
BEGIN
    ALTER TABLE dbo.CON_TipoCambio
        ADD CONSTRAINT UQ_CON_TipoCambio_Cuenta_Fecha_Moneda
            UNIQUE (IdCuentaAdministradora, Fecha, IdMoneda);
END;
