-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Lista los centros de costo configurados por empresa con filtro y paginacion.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarCentroCostoConfiguracionEmpresa
    @IdEmpresa INT,
    @SoloActivos BIT = 1,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                c.IdCentroCostoConfiguracionEmpresa AS IdCentroCosto,
                c.IdEmpresa,
                c.Codigo AS CodigoCentroCosto,
                c.Nombre AS NombreCentroCosto,
                c.Estado
            FROM dbo.CON_CentroCostoConfiguracionEmpresa AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND (@SoloActivos = 0 OR c.Estado = 1)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR c.Codigo LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.Nombre LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdCentroCosto,
            b.IdEmpresa,
            b.CodigoCentroCosto,
            b.NombreCentroCosto,
            b.Estado,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.CodigoCentroCosto ASC
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
