-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Cabecera del proceso de ajuste de cuentas por empresa y periodo.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Registra un unico proceso AJU por periodo para permitir regeneraciones del ajuste de cuentas sobre cuentas de analisis.

IF OBJECT_ID(N'dbo.CON_AjusteCuentaProceso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AjusteCuentaProceso
    (
        IdAjusteCuentaProceso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AjusteCuentaProceso PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Periodo CHAR(6) NOT NULL,
        IdOrigen INT NOT NULL,
        FechaAsiento DATE NOT NULL,
        TotalCuentas INT NOT NULL CONSTRAINT DF_CON_AjusteCuentaProceso_TotalCuentas DEFAULT (0),
        TotalAsientos INT NOT NULL CONSTRAINT DF_CON_AjusteCuentaProceso_TotalAsientos DEFAULT (0),
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProceso_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProceso_TotalHaber DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProceso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_AjusteCuentaProceso
        ADD CONSTRAINT FK_CON_AjusteCuentaProceso_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_AjusteCuentaProceso
        ADD CONSTRAINT FK_CON_AjusteCuentaProceso_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_AjusteCuentaProceso
        ADD CONSTRAINT CK_CON_AjusteCuentaProceso_Periodo
            CHECK (
                Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
                AND RIGHT(Periodo, 2) BETWEEN '01' AND '12'
            );

    ALTER TABLE dbo.CON_AjusteCuentaProceso
        ADD CONSTRAINT UQ_CON_AjusteCuentaProceso_IdEmpresa_Periodo
            UNIQUE (IdEmpresa, Periodo);
END;
