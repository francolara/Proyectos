-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Carga Factura 01 y Boleta 03 sin duplicar ni
--                sobrescribir registros existentes del catalogo interno.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    INSERT INTO dbo.TipoDocumento (Codigo, Nombre)
    SELECT origen.Codigo, origen.Nombre
    FROM (VALUES (N'01', N'Factura'), (N'03', N'Boleta'))
        AS origen (Codigo, Nombre)
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM dbo.TipoDocumento AS tipo WITH (UPDLOCK, HOLDLOCK)
        WHERE tipo.Codigo = origen.Codigo
    );

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO
