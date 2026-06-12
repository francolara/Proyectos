GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 11/06/2026 | Devuelve el resumen agregado del modulo de boletines deportivos para el panel super admin sin cargar el listado completo.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_AdminResumen
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            COUNT(1) AS TotalBoletines,
            SUM(CASE WHEN b.Activo = 1 THEN 1 ELSE 0 END) AS TotalActivos,
            SUM(CASE WHEN b.Activo = 0 THEN 1 ELSE 0 END) AS TotalInactivos,
            SUM(CASE WHEN b.TipoRegistro = 'U' THEN 1 ELSE 0 END) AS TotalUsuarios,
            SUM(CASE WHEN b.TipoRegistro = 'A' THEN 1 ELSE 0 END) AS TotalPlataforma
        FROM dbo.BoletinesDeportivos b;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
