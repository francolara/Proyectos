-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega cargo y sueldo base mensual en soles a la ficha del trabajador.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'CargoId') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD CargoId BIGINT NULL CONSTRAINT FK_TrabajadorHYT_Cargo REFERENCES dbo.CargoTrabajadorHYT(CargoId);
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'SueldoMensual') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD SueldoMensual DECIMAL(18,2) NULL CONSTRAINT CK_TrabajadorHYT_Sueldo CHECK (SueldoMensual > 0);
GO
