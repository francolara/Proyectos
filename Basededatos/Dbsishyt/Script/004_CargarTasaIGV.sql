-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Inicializa 0.18 desde la fecha de configuracion del proyecto,
--                solo si no existen tasas; no establece vigencia legal.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS (SELECT 1 FROM dbo.TasaIGV AS tasa WITH (UPDLOCK, HOLDLOCK))
    BEGIN
        INSERT INTO dbo.TasaIGV (FechaDesde, FechaHasta, Valor)
        VALUES (CONVERT(DATE, '20261103', 112), NULL, 0.18);
    END;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO
