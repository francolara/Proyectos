-- Firma: FRANCO LARA | Create date: 10/11/2026 | Maestro de conceptos de planilla con valor opcional.
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.ConceptoPlanillaHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConceptoPlanillaHYT
    (
        ConceptoId BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ConceptoPlanillaHYT PRIMARY KEY,
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT UQ_ConceptoPlanillaHYT_Nombre UNIQUE,
        Naturaleza VARCHAR(5) NOT NULL CHECK (Naturaleza IN ('CARGO','ABONO')),
        Calculo VARCHAR(10) NOT NULL CHECK (Calculo IN ('IMPORTE','PORCENTAJE')),
        Valor DECIMAL(20,4) NULL,
        Regimen VARCHAR(3) NULL CHECK (Regimen IN ('AFP','ONP')),
        Activo BIT NOT NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_ConceptoPlanillaHYT_Valor CHECK (Valor >= 0 AND
            ((Calculo='PORCENTAJE' AND Valor<=100) OR (Calculo='IMPORTE' AND Valor=ROUND(Valor,2)))),
        CONSTRAINT CK_ConceptoPlanillaHYT_Regimen CHECK (Regimen IS NULL OR Naturaleza='CARGO')
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.ConceptoPlanillaHYT') AND name=N'UX_ConceptoPlanillaHYT_Regimen')
    CREATE UNIQUE INDEX UX_ConceptoPlanillaHYT_Regimen ON dbo.ConceptoPlanillaHYT(Regimen) WHERE Regimen IS NOT NULL;
GO
