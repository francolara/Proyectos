-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/07/2026
-- Description:   Cabecera del proceso de diferencia en cambio por empresa y periodo.
-- =============================================
-- Firma: FRANCO LARA - 01/07/2026 | Registra un unico proceso por periodo para controlar regeneraciones de diferencia en cambio y el tipo de cambio aplicado al cierre.

IF OBJECT_ID(N'dbo.CON_DiferenciaCambioProceso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_DiferenciaCambioProceso
    (
        IdDiferenciaCambioProceso INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_DiferenciaCambioProceso PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Periodo CHAR(6) NOT NULL,
        IdOrigen INT NOT NULL,
        FechaAsiento DATE NOT NULL,
        UsaTipoCambioSbs BIT NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_UsaTipoCambioSbs DEFAULT (0),
        TipoCambioCompra DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_TipoCambioCompra DEFAULT (0),
        TipoCambioVenta DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_TipoCambioVenta DEFAULT (0),
        TotalCuentas INT NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_TotalCuentas DEFAULT (0),
        TotalAsientos INT NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_TotalAsientos DEFAULT (0),
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_TotalHaber DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProceso_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_DiferenciaCambioProceso
        ADD CONSTRAINT FK_CON_DiferenciaCambioProceso_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_DiferenciaCambioProceso
        ADD CONSTRAINT FK_CON_DiferenciaCambioProceso_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_DiferenciaCambioProceso
        ADD CONSTRAINT CK_CON_DiferenciaCambioProceso_Periodo
            CHECK (
                Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
                AND RIGHT(Periodo, 2) BETWEEN '01' AND '12'
            );

    ALTER TABLE dbo.CON_DiferenciaCambioProceso
        ADD CONSTRAINT UQ_CON_DiferenciaCambioProceso_IdEmpresa_Periodo
            UNIQUE (IdEmpresa, Periodo);
END;
