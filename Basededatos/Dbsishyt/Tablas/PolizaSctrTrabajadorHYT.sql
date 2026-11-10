-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   PolizaSctrTrabajadorHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PolizaSctrTrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PolizaSctrTrabajadorHYT
    (
        PolizaId BIGINT NOT NULL CONSTRAINT FK_PolizaSctrTrabajadorHYT_Poliza REFERENCES dbo.PolizaSctrHYT(PolizaId),
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_PolizaSctrTrabajadorHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        CONSTRAINT PK_PolizaSctrTrabajadorHYT PRIMARY KEY (PolizaId, TrabajadorId)
    );
END;
GO
