-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Crea la tabla de control de estado del periodo contable por empresa.
-- =============================================
-- Firma: FRANCO LARA - 30/06/2026 | Crea la tabla CON_PeriodoContableEstado para abrir o cerrar periodos contables por empresa y bloquear la operativa de registros.

IF OBJECT_ID(N'dbo.CON_PeriodoContableEstado', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_PeriodoContableEstado
    (
        IdPeriodoContableEstado INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_PeriodoContableEstado PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Periodo CHAR(6) NOT NULL,
        Cerrado BIT NOT NULL CONSTRAINT DF_CON_PeriodoContableEstado_Cerrado DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_PeriodoContableEstado_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL,
        FechaCierre DATETIME2(0) NULL,
        UsuarioCierre NVARCHAR(450) NULL,
        FechaApertura DATETIME2(0) NULL,
        UsuarioApertura NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_PeriodoContableEstado
        ADD CONSTRAINT FK_CON_PeriodoContableEstado_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_PeriodoContableEstado
        ADD CONSTRAINT UQ_CON_PeriodoContableEstado_Empresa_Periodo
            UNIQUE (IdEmpresa, Periodo);

    ALTER TABLE dbo.CON_PeriodoContableEstado
        ADD CONSTRAINT CK_CON_PeriodoContableEstado_Periodo
            CHECK (Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]');
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_CON_PeriodoContableEstado_SEG_Empresa'
)
BEGIN
    ALTER TABLE dbo.CON_PeriodoContableEstado
        ADD CONSTRAINT FK_CON_PeriodoContableEstado_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.key_constraints
    WHERE name = N'UQ_CON_PeriodoContableEstado_Empresa_Periodo'
)
BEGIN
    ALTER TABLE dbo.CON_PeriodoContableEstado
        ADD CONSTRAINT UQ_CON_PeriodoContableEstado_Empresa_Periodo
            UNIQUE (IdEmpresa, Periodo);
END;

