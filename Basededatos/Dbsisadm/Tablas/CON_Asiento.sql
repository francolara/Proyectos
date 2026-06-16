-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Cabecera de asientos contables por empresa con correlativo por origen y periodo mensual.
-- =============================================

IF OBJECT_ID(N'dbo.CON_Asiento', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_Asiento
    (
        IdAsiento INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_Asiento PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdOrigen INT NOT NULL,
        Ejercicio SMALLINT NOT NULL,
        Mes TINYINT NOT NULL,
        Periodo CHAR(6) NOT NULL,
        NumeroAsiento INT NOT NULL,
        FechaAsiento DATE NOT NULL,
        Glosa NVARCHAR(500) NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_Asiento_TipoCambio DEFAULT (1),
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_Asiento_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_Asiento_TotalHaber DEFAULT (0),
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_CON_Asiento_Estado DEFAULT (N'BORRADOR'),
        ReferenciaExterna NVARCHAR(100) NULL,
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_Asiento_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT FK_CON_Asiento_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT FK_CON_Asiento_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT FK_CON_Asiento_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT CK_CON_Asiento_Mes
            CHECK (Mes BETWEEN 1 AND 12);

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT CK_CON_Asiento_Periodo
            CHECK (
                Periodo = CONVERT(CHAR(4), Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), Mes), 2)
            );

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT CK_CON_Asiento_Totales
            CHECK (TotalDebe >= 0 AND TotalHaber >= 0);

    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT UQ_CON_Asiento_Numero
            UNIQUE (IdEmpresa, IdOrigen, Periodo, NumeroAsiento);
END;
