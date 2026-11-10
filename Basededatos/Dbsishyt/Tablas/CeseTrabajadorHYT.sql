-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.CeseTrabajadorHYT para las etapas 3 y 4 de Planillas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Agrega datos de reingreso e idempotencia.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.CeseTrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CeseTrabajadorHYT
    (
        CeseId BIGINT IDENTITY NOT NULL CONSTRAINT PK_CeseTrabajadorHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_CeseTrabajadorHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        FechaCese DATE NOT NULL,
        UltimoDiaTrabajado DATE NOT NULL,
        Motivo NVARCHAR(300) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_CeseTrabajadorHYT_Operacion UNIQUE,
        CONSTRAINT CK_CeseTrabajadorHYT_Fechas CHECK (UltimoDiaTrabajado <= FechaCese),
        CONSTRAINT UQ_CeseTrabajadorHYT_TrabajadorFecha UNIQUE (TrabajadorId, FechaCese)
    );
END;
GO

IF COL_LENGTH(N'dbo.CeseTrabajadorHYT', N'FechaReingreso') IS NULL
BEGIN
    ALTER TABLE dbo.CeseTrabajadorHYT ADD
        FechaReingreso DATE NULL,
        MotivoReingreso NVARCHAR(300) NULL,
        OperacionReingreso UNIQUEIDENTIFIER NULL;
END;
GO
IF OBJECT_ID(N'dbo.CK_CeseTrabajadorHYT_Reingreso', N'C') IS NULL
BEGIN
    ALTER TABLE dbo.CeseTrabajadorHYT ADD CONSTRAINT CK_CeseTrabajadorHYT_Reingreso CHECK (
        (FechaReingreso IS NULL AND MotivoReingreso IS NULL AND OperacionReingreso IS NULL)
        OR (FechaReingreso IS NOT NULL AND FechaReingreso > FechaCese AND OperacionReingreso IS NOT NULL));
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.CeseTrabajadorHYT')
               AND name = N'UX_CeseTrabajadorHYT_OperacionReingreso')
    CREATE UNIQUE INDEX UX_CeseTrabajadorHYT_OperacionReingreso
    ON dbo.CeseTrabajadorHYT (OperacionReingreso) WHERE OperacionReingreso IS NOT NULL;
GO
