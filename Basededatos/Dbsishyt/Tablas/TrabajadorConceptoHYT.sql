-- Firma: FRANCO LARA | Create date: 10/11/2026 | Relaciona los conceptos habituales con cada trabajador.
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.TrabajadorConceptoHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TrabajadorConceptoHYT
    (
        TrabajadorId BIGINT NOT NULL REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        ConceptoId BIGINT NOT NULL REFERENCES dbo.ConceptoPlanillaHYT(ConceptoId),
        CONSTRAINT PK_TrabajadorConceptoHYT PRIMARY KEY (TrabajadorId,ConceptoId)
    );
END;
GO
