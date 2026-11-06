-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define NumeracionInterna para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.NumeracionInterna', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NumeracionInterna
    (
        TipoRegistro NVARCHAR(40) NOT NULL CONSTRAINT PK_NumeracionInterna PRIMARY KEY,
        Prefijo NVARCHAR(20) NOT NULL,
        UltimoNumero BIGINT NOT NULL DEFAULT 0,
        CONSTRAINT CK_Numeracion_Numero CHECK (UltimoNumero >= 0)
    );
END;
GO
