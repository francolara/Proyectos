-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.JornadaHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.JornadaHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.JornadaHYT
    (
        JornadaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_JornadaHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_Jornada_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        Entrada DATETIME2(0) NOT NULL,
        Salida DATETIME2(0) NULL,
        MinutosExtrasAprobados INT NOT NULL CONSTRAINT DF_Jornada_Extras DEFAULT (0),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT UQ_Jornada_TrabajadorEntrada UNIQUE (TrabajadorId, Entrada),
        CONSTRAINT CK_Jornada_Horario CHECK (MinutosExtrasAprobados >= 0 AND
            ((Salida IS NULL AND MinutosExtrasAprobados = 0) OR
             (Salida IS NOT NULL AND Salida > Entrada AND MinutosExtrasAprobados <= DATEDIFF(MINUTE, Entrada, Salida))))
    );
END;
GO
