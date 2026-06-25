-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Lista parametros por empresa con busqueda y paginacion server-side.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarParametrosPorEmpresa
    @IdEmpresa INT,
    @TipoParametro VARCHAR(30) = NULL,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TipoParametroTrabajo VARCHAR(30) = NULLIF(LTRIM(RTRIM(@TipoParametro)), '')
        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                pe.IdParametroEmpresa,
                pe.IdEmpresa,
                pe.TipoParametro,
                pe.CodigoParametro,
                pe.ValorParametro,
                pe.DescripcionParametro,
                pe.FecIni,
                pe.FecFin,
                pe.Activo
            FROM dbo.ADM_ParametroEmpresa AS pe
            WHERE pe.IdEmpresa = @IdEmpresa
              AND (@TipoParametroTrabajo IS NULL OR pe.TipoParametro = @TipoParametroTrabajo)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR pe.CodigoParametro LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR pe.ValorParametro LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR pe.DescripcionParametro LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdParametroEmpresa,
            b.IdEmpresa,
            b.TipoParametro,
            b.CodigoParametro,
            b.ValorParametro,
            b.DescripcionParametro,
            b.FecIni,
            b.FecFin,
            b.Activo,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.TipoParametro ASC, b.CodigoParametro ASC
        OFFSET CASE WHEN @NumeroPaginaTrabajo IS NULL OR @TamanoPaginaTrabajo IS NULL THEN 0 ELSE (@NumeroPaginaTrabajo - 1) * @TamanoPaginaTrabajo END ROWS
        FETCH NEXT CASE WHEN @NumeroPaginaTrabajo IS NULL OR @TamanoPaginaTrabajo IS NULL THEN 2147483647 ELSE @TamanoPaginaTrabajo END ROWS ONLY;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
