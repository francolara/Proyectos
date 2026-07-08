-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Cabecera del proceso de asiento de cierre anual por empresa y ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Registra un unico proceso anual de cierre para controlar la regeneracion de ganancias/perdidas e inventarios con el TC de 31/12.

IF OBJECT_ID(N'dbo.CON_CierreProceso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CierreProceso
    (
        IdCierreProceso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CierreProceso PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Anio SMALLINT NOT NULL,
        IdOrigen INT NOT NULL,
        FechaAsiento DATE NOT NULL,
        UsaTipoCambioSbs BIT NOT NULL CONSTRAINT DF_CON_CierreProceso_UsaTipoCambioSbs DEFAULT (0),
        TipoCambioCompra DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_CierreProceso_TipoCambioCompra DEFAULT (0),
        TipoCambioVenta DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_CierreProceso_TipoCambioVenta DEFAULT (0),
        ProcesaGananciasPerdidas BIT NOT NULL CONSTRAINT DF_CON_CierreProceso_ProcesaGananciasPerdidas DEFAULT (0),
        ProcesaInventarios BIT NOT NULL CONSTRAINT DF_CON_CierreProceso_ProcesaInventarios DEFAULT (0),
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
        ADD CONSTRAINT CK_CON_CierreProceso_Anio
            CHECK (Anio BETWEEN 2000 AND 9999);

    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT UQ_CON_CierreProceso_IdEmpresa_Anio
            UNIQUE (IdEmpresa, Anio);
END;
