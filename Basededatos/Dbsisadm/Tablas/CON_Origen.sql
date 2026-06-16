-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Origenes contables por empresa para clasificar asientos.
-- =============================================

IF OBJECT_ID(N'dbo.CON_Origen', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_Origen
    (
        IdOrigen INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_Origen PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        CodigoOrigen VARCHAR(10) NOT NULL,
        NombreOrigen NVARCHAR(150) NOT NULL,
        ModuloOrigen NVARCHAR(50) NOT NULL,
        PermiteRegistroManual BIT NOT NULL CONSTRAINT DF_CON_Origen_PermiteRegistroManual DEFAULT (1),
        Estado BIT NOT NULL CONSTRAINT DF_CON_Origen_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_Origen_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_Origen
        ADD CONSTRAINT FK_CON_Origen_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_Origen
        ADD CONSTRAINT UQ_CON_Origen_IdEmpresa_CodigoOrigen
            UNIQUE (IdEmpresa, CodigoOrigen);
END;
