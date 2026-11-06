-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Carga monedas PEN y USD sin duplicados ni sobrescrituras.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    INSERT INTO dbo.Moneda (Codigo, Nombre, Simbolo)
    SELECT origen.Codigo, origen.Nombre, origen.Simbolo
    FROM (VALUES (N'PEN', N'Soles', N'S/'), (N'USD', N'Dólares', N'$')) AS origen (Codigo, Nombre, Simbolo)
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Moneda AS moneda WITH (UPDLOCK, HOLDLOCK) WHERE moneda.Codigo = origen.Codigo);
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO
