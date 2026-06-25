-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Crea la tabla maestra de centros de costo contables.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CentroCosto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CentroCosto
    (
        IdCentroCosto INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CentroCosto PRIMARY KEY,
        Codigo VARCHAR(20) NOT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_CentroCosto_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_CentroCosto
        ADD CONSTRAINT UQ_CON_CentroCosto_Codigo UNIQUE (Codigo);
END;
