-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Crea la cabecera de movimientos de caja y bancos por empresa y cuenta corriente con correlativo interno por periodo.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Amplia la cabecera de movimientos bancarios agregando TipoCambio y Observacion para el registro operativo.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Vincula cada movimiento bancario con el asiento contable generado automaticamente.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega el enlace funcional para transferencias entre cuentas guardando codigo comun, rol emisor/receptor y movimiento relacionado.
-- =============================================

IF OBJECT_ID(N'dbo.BAN_MovimientoBanco', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BAN_MovimientoBanco
    (
        IdMovimientoBanco INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_BAN_MovimientoBanco PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdBancoConfiguracionEmpresa INT NOT NULL,
        TipoMovimiento CHAR(1) NOT NULL,
        IdOpeBancaria CHAR(2) NOT NULL,
        FechaEmision DATE NOT NULL,
        TipoCambio DECIMAL(18, 6) NOT NULL CONSTRAINT DF_BAN_MovimientoBanco_TipoCambio DEFAULT (1),
        NumeroMovimiento INT NOT NULL,
        IdAsiento INT NULL,
        IdTransferenciaCuenta UNIQUEIDENTIFIER NULL,
        RolTransferencia CHAR(1) NULL,
        IdMovimientoBancoRelacionado INT NULL,
        IdPersona INT NULL,
        NumeroDocumento VARCHAR(20) NULL,
        Glosa NVARCHAR(300) NOT NULL,
        Observacion NVARCHAR(500) NULL,
        ImporteTotal DECIMAL(18, 2) NOT NULL CONSTRAINT DF_BAN_MovimientoBanco_ImporteTotal DEFAULT (0),
        Activo BIT NOT NULL CONSTRAINT DF_BAN_MovimientoBanco_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_BAN_MovimientoBanco_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_CON_BancosConfiguracionEmpresa
            FOREIGN KEY (IdBancoConfiguracionEmpresa) REFERENCES dbo.CON_BancosConfiguracionEmpresa (IdBancoConfiguracionEmpresa);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_ADM_Persona
            FOREIGN KEY (IdPersona) REFERENCES dbo.ADM_Persona (IdPersona);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_MovimientoRelacionado
            FOREIGN KEY (IdMovimientoBancoRelacionado) REFERENCES dbo.BAN_MovimientoBanco (IdMovimientoBanco);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_TipoMovimiento
            CHECK (TipoMovimiento IN ('I', 'E'));

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_TipoCambio
            CHECK (TipoCambio > 0);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_ImporteTotal
            CHECK (ImporteTotal >= 0);

    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_RolTransferencia
            CHECK (RolTransferencia IS NULL OR RolTransferencia IN ('E', 'I'));
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'IdAsiento') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD IdAsiento INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_BAN_MovimientoBanco_CON_Asiento'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'IdTransferenciaCuenta') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD IdTransferenciaCuenta UNIQUEIDENTIFIER NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'RolTransferencia') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD RolTransferencia CHAR(1) NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'IdMovimientoBancoRelacionado') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD IdMovimientoBancoRelacionado INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_BAN_MovimientoBanco_RolTransferencia'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_RolTransferencia
            CHECK (RolTransferencia IS NULL OR RolTransferencia IN ('E', 'I'));
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_BAN_MovimientoBanco_MovimientoRelacionado'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_MovimientoRelacionado
            FOREIGN KEY (IdMovimientoBancoRelacionado) REFERENCES dbo.BAN_MovimientoBanco (IdMovimientoBanco);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'IX_BAN_MovimientoBanco_IdTransferenciaCuenta'
      AND object_id = OBJECT_ID(N'dbo.BAN_MovimientoBanco')
)
BEGIN
    CREATE INDEX IX_BAN_MovimientoBanco_IdTransferenciaCuenta
        ON dbo.BAN_MovimientoBanco (IdTransferenciaCuenta, RolTransferencia);
END;
