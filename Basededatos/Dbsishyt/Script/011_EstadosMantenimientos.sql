-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Baja lógica de datos HYT por cliente y partidas del tarifario.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ProveedorCliente', N'Activo') IS NULL
    ALTER TABLE dbo.ProveedorCliente ADD Activo BIT NOT NULL CONSTRAINT DF_ProveedorCliente_Activo DEFAULT (1) WITH VALUES;
IF COL_LENGTH(N'dbo.TarifarioPartida', N'Activo') IS NULL
    ALTER TABLE dbo.TarifarioPartida ADD Activo BIT NOT NULL CONSTRAINT DF_TarifarioPartida_Activo DEFAULT (1) WITH VALUES;
GO
