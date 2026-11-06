-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define tasas de IGV por vigencia sin alterar historicos.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TasaIGV', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TasaIGV
    (
        FechaDesde DATE NOT NULL,
        FechaHasta DATE NULL,
        Valor DECIMAL(9,6) NOT NULL,
        CONSTRAINT PK_TasaIGV PRIMARY KEY (FechaDesde),
        CONSTRAINT UQ_TasaIGV_Valor UNIQUE (FechaDesde, Valor),
        CONSTRAINT CK_TasaIGV_Valor CHECK (Valor >= 0 AND Valor <= 1),
        CONSTRAINT CK_TasaIGV_Vigencia CHECK (FechaHasta IS NULL OR FechaHasta > FechaDesde)
    );
END;
GO
