-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Inicializa correlativos internos sin alterar los existentes.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    INSERT INTO dbo.NumeracionInterna (TipoRegistro, Prefijo, UltimoNumero)
    SELECT origen.TipoRegistro, origen.Prefijo, 0
    FROM (VALUES (N'PRESUPUESTO', N'PPTO-'), (N'OC_CLIENTE', N'OC-')) AS origen (TipoRegistro, Prefijo)
    WHERE NOT EXISTS
    (
        SELECT 1 FROM dbo.NumeracionInterna AS numeracion WITH (UPDLOCK, HOLDLOCK)
        WHERE numeracion.TipoRegistro = origen.TipoRegistro
    );
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO
