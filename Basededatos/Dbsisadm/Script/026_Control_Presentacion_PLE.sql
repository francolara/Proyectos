-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Migra el historial PLE para controlar presentacion y snapshots incrementales del plan.
-- =============================================

SET NOCOUNT ON;

BEGIN TRY
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'CodigoFormatoComplementario') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD CodigoFormatoComplementario VARCHAR(10) NULL;
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'NombreArchivoComplementario') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD NombreArchivoComplementario NVARCHAR(250) NULL;
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'CantidadRegistrosComplementario') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD CantidadRegistrosComplementario INT NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_CantidadComplementaria DEFAULT (0);
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'HuellaPlanContable') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD HuellaPlanContable CHAR(64) NULL;
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'PlanContableSnapshot') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD PlanContableSnapshot NVARCHAR(MAX) NULL;
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'PlanPresentado') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD PlanPresentado BIT NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_PlanPresentado DEFAULT (0);
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'FechaPresentacion') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD FechaPresentacion DATETIME2(0) NULL;
    IF COL_LENGTH(N'dbo.CON_LibroElectronicoGeneracion', N'UsuarioPresentacion') IS NULL
        ALTER TABLE dbo.CON_LibroElectronicoGeneracion ADD UsuarioPresentacion NVARCHAR(450) NULL;
END TRY
BEGIN CATCH
    DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
    SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH
