-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista motivos de traslado por documento para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantMotivosTraslado_Listar
    @IdDocumento VARCHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT MT.IdMotivoTraslado, MT.GlsMotivoTraslado,
               CASE WHEN MT.IndGeneraVale = 'S'
                    THEN CASE WHEN MT.IndTipoVale = 'I' THEN 'INGRESO' ELSE 'SALIDA' END
                    ELSE '---' END AS IndTipoVale,
               CASE WHEN MT.IndGeneraVale = 'S' THEN 'SI' ELSE 'NO' END AS IndGeneraVale
        FROM dbo.MotivosTraslados AS MT
        WHERE MT.IdDocumento = @IdDocumento
          AND (@Busqueda = '' OR MT.GlsMotivoTraslado LIKE '%' + @Busqueda + '%')
        ORDER BY MT.IdMotivoTraslado;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
