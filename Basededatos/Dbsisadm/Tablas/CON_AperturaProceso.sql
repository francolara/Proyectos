-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Cabecera del proceso de asiento de apertura por empresa y ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Registra un unico proceso de apertura por empresa y anio, con el periodo de saldos usado y el asiento generado en 00.

IF OBJECT_ID(N'dbo.CON_AperturaProceso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AperturaProceso
    (
        IdAperturaProceso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AperturaProceso PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        AnioApertura SMALLINT NOT NULL,
        AnioSaldo SMALLINT NOT NULL,
        MesSaldoHasta TINYINT NOT NULL,
        PeriodoSaldoHasta CHAR(6) NOT NULL,
        IdOrigen INT NOT NULL,
        FechaAsiento DATE NOT NULL,
        UsaTipoCambioSbs BIT NOT NULL CONSTRAINT DF_CON_AperturaProceso_UsaTipoCambioSbs DEFAULT (0),
        TipoCambioCompra DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_AperturaProceso_TipoCambioCompra DEFAULT (0),
        TipoCambioVenta DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_AperturaProceso_TipoCambioVenta DEFAULT (0),
        IdAsiento INT NULL,
        NumeroAsiento INT NULL,
        TotalLineas INT NOT NULL CONSTRAINT DF_CON_AperturaProceso_TotalLineas DEFAULT (0),
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AperturaProceso_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AperturaProceso_TotalHaber DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_AperturaProceso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT FK_CON_AperturaProceso_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT FK_CON_AperturaProceso_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT FK_CON_AperturaProceso_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT CK_CON_AperturaProceso_Anio
            CHECK (AnioApertura BETWEEN 2000 AND 9999 AND AnioSaldo BETWEEN 1999 AND 9998);

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT CK_CON_AperturaProceso_MesSaldo
            CHECK (MesSaldoHasta BETWEEN 0 AND 15);

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT CK_CON_AperturaProceso_PeriodoSaldo
            CHECK (
                PeriodoSaldoHasta LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
                AND RIGHT(PeriodoSaldoHasta, 2) BETWEEN '00' AND '15'
            );

    ALTER TABLE dbo.CON_AperturaProceso
        ADD CONSTRAINT UQ_CON_AperturaProceso_IdEmpresa_AnioApertura
            UNIQUE (IdEmpresa, AnioApertura);
END;
