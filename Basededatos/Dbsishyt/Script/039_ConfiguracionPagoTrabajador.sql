-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega horario habitual, regimen previsional y frecuencia de pago.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'HoraEntradaHabitual') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD HoraEntradaHabitual TIME(0) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'HoraSalidaHabitual') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD HoraSalidaHabitual TIME(0) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'RegimenPrevisional') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD RegimenPrevisional VARCHAR(3) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'AdministradoraAFP') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD AdministradoraAFP NVARCHAR(100) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'FrecuenciaPago') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD FrecuenciaPago VARCHAR(10) NULL;
GO
IF OBJECT_ID(N'dbo.CK_TrabajadorHYT_Planificacion', N'C') IS NULL
    ALTER TABLE dbo.TrabajadorHYT WITH CHECK ADD CONSTRAINT CK_TrabajadorHYT_Planificacion CHECK (
        ((HoraEntradaHabitual IS NULL AND HoraSalidaHabitual IS NULL)
         OR (HoraEntradaHabitual IS NOT NULL AND HoraSalidaHabitual IS NOT NULL AND HoraEntradaHabitual <> HoraSalidaHabitual))
        AND (RegimenPrevisional IS NULL OR RegimenPrevisional IN ('AFP', 'ONP'))
        AND (AdministradoraAFP IS NULL OR (RegimenPrevisional IS NOT NULL AND RegimenPrevisional = 'AFP'))
        AND (FrecuenciaPago IS NULL OR FrecuenciaPago IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')));
GO
