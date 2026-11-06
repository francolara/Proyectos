-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Carga los tratamientos CON IGV y SIN IGV
--                sin duplicar ni sobrescribir registros existentes.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    INSERT INTO dbo.TipoIGV (Codigo, Nombre, IncluyeIGV)
    SELECT origen.Codigo, origen.Nombre, origen.IncluyeIGV
    FROM (VALUES
        (N'CON_IGV', N'CON IGV', CAST(1 AS BIT)),
        (N'SIN_IGV', N'SIN IGV', CAST(0 AS BIT))
    ) AS origen (Codigo, Nombre, IncluyeIGV)
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM dbo.TipoIGV AS tipo WITH (UPDLOCK, HOLDLOCK)
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
