-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista tipos de datos para el mantenimiento general.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantDatos_Tipos_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT DC.DescDatos, DC.IdTipoDatos
        FROM dbo.DatosCab AS DC
        ORDER BY DC.DescDatos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
