-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.ObligacionPersonalHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.ObligacionPersonalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ObligacionPersonalHYT
    (
        ObligacionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ObligacionPersonalHYT PRIMARY KEY,
        Tipo VARCHAR(3) NOT NULL CONSTRAINT CK_ObligacionPersonalHYT_Tipo CHECK (Tipo IN ('AFP', 'ONP')),
        Entidad NVARCHAR(120) NOT NULL,
        Mes DATE NOT NULL CONSTRAINT CK_ObligacionPersonalHYT_Mes CHECK (DAY(Mes) = 1),
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_ObligacionPersonalHYT_Importe CHECK (Importe > 0),
        Vencimiento DATE NOT NULL,
        FechaPago DATE NULL,
        ReferenciaPago NVARCHAR(100) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT UQ_ObligacionPersonalHYT_TipoEntidadMes UNIQUE (Tipo, Entidad, Mes),
        CONSTRAINT CK_ObligacionPersonalHYT_Pago CHECK (
            (FechaPago IS NULL AND ReferenciaPago IS NULL) OR
            (FechaPago IS NOT NULL AND ReferenciaPago IS NOT NULL AND LEN(ReferenciaPago) > 0))
    );
END;
GO
