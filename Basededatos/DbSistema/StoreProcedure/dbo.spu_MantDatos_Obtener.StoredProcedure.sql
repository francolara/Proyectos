-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Obtiene un registro del mantenimiento general de datos.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantDatos_Obtener
    @IdTipoDatos CHAR(3),
    @IdDato CHAR(5)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT D.IdDato, D.IdTipoDatos, D.GlsDato, D.EstDato, D.Abreviatura,
               D.IdSunat, D.IdCtaVtaSoles, D.IdCtaVtaDolares,
               D.IdCtaVtaTerSoles, D.IdCtaVtaTerDolares, D.Destino,
               D.Orden, D.IdCtaCorriente
        FROM dbo.Datos AS D
        WHERE D.IdTipoDatos = @IdTipoDatos AND D.IdDato = @IdDato;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
