-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista centros de costo para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantCentroCosto_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            CC.IdCentroCosto,
            CC.GlsCentroCosto
        FROM dbo.CentrosCosto AS CC
        WHERE CC.IdEmpresa = @IdEmpresa
          AND (
                @Busqueda = ''
                OR CC.IdCentroCosto LIKE '%' + @Busqueda + '%'
                OR CC.GlsCentroCosto LIKE '%' + @Busqueda + '%'
              )
        ORDER BY CC.IdCentroCosto;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
