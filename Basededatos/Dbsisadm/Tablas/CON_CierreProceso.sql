-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Cabecera del proceso de asiento de cierre anual por empresa y ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Registra un unico proceso anual de cierre para controlar la regeneracion de ganancias/perdidas e inventarios con el TC de 31/12.
-- Firma: FRANCO LARA - 13/08/2026 | Adapta el proceso para registrar un unico asiento compuesto, su periodo de corte, periodo de generacion, correlativo y total de lineas.
-- Firma: FRANCO LARA - 22/08/2026 | Fija el cierre de Inventario en el periodo 14 y restringe el corte de saldos a los periodos 00-13.

IF OBJECT_ID(N'dbo.CON_CierreProceso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CierreProceso
    (
        IdCierreProceso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CierreProceso PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Anio SMALLINT NOT NULL,
        MesSaldoHasta TINYINT NOT NULL CONSTRAINT DF_CON_CierreProceso_MesSaldoHasta DEFAULT (13),
        PeriodoSaldoHasta CHAR(6) NOT NULL,
        MesGeneracion TINYINT NOT NULL CONSTRAINT DF_CON_CierreProceso_MesGeneracion DEFAULT (14),
        PeriodoGeneracion CHAR(6) NOT NULL,
        IdOrigen INT NOT NULL,
        FechaAsiento DATE NOT NULL,
        UsaTipoCambioSbs BIT NOT NULL CONSTRAINT DF_CON_CierreProceso_UsaTipoCambioSbs DEFAULT (0),
        TipoCambioCompra DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_CierreProceso_TipoCambioCompra DEFAULT (0),
        TipoCambioVenta DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_CierreProceso_TipoCambioVenta DEFAULT (0),
        ProcesaGananciasPerdidas BIT NOT NULL CONSTRAINT DF_CON_CierreProceso_ProcesaGananciasPerdidas DEFAULT (0),
        ProcesaInventarios BIT NOT NULL CONSTRAINT DF_CON_CierreProceso_ProcesaInventarios DEFAULT (0),
        IdAsiento INT NULL,
        NumeroAsiento INT NULL,
        TotalLineas INT NOT NULL CONSTRAINT DF_CON_CierreProceso_TotalLineas DEFAULT (0),
        TotalCuentas INT NOT NULL CONSTRAINT DF_CON_CierreProceso_TotalCuentas DEFAULT (0),
        TotalAsientos INT NOT NULL CONSTRAINT DF_CON_CierreProceso_TotalAsientos DEFAULT (0),
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProceso_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProceso_TotalHaber DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CierreProceso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT FK_CON_CierreProceso_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT FK_CON_CierreProceso_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT FK_CON_CierreProceso_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT CK_CON_CierreProceso_Anio
            CHECK (Anio BETWEEN 2000 AND 9999);

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT CK_CON_CierreProceso_Meses
            CHECK (MesSaldoHasta BETWEEN 0 AND 13 AND MesGeneracion = 14);

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT CK_CON_CierreProceso_Periodos
            CHECK (
                PeriodoSaldoHasta = CONVERT(CHAR(4), Anio) + RIGHT('0' + CONVERT(VARCHAR(2), MesSaldoHasta), 2)
                AND PeriodoGeneracion = CONVERT(CHAR(4), Anio) + RIGHT('0' + CONVERT(VARCHAR(2), MesGeneracion), 2)
            );

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT UQ_CON_CierreProceso_IdEmpresa_Anio
            UNIQUE (IdEmpresa, Anio);
END;
