-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista registros del tipo seleccionado en el mantenimiento de datos.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantDatos_Listar
    @IdTipoDatos CHAR(3)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT D.IdDato, D.IdTipoDatos, D.GlsDato, D.EstDato, D.Abreviatura,
               D.IdSunat, D.IdCtaVtaSoles, D.IdCtaVtaDolares,
               D.IdCtaVtaTerSoles, D.IdCtaVtaTerDolares, D.Destino,
               D.Orden, D.IdCtaCorriente
        FROM dbo.Datos AS D
        WHERE D.IdTipoDatos = @IdTipoDatos
        ORDER BY D.IdDato;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
