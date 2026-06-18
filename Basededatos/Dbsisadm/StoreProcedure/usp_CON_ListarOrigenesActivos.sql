-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista los origenes contables activos de una empresa.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Ajusta el mantenimiento web para listar origenes con busqueda y paginacion server-side segun filtro.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarOrigenesActivos
    @IdEmpresa INT,
    @SoloActivos BIT = 1,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                o.IdOrigen,
                o.CodigoOrigen,
                o.NombreOrigen,
                o.ModuloOrigen,
                o.PermiteRegistroManual,
                o.Estado
            FROM dbo.CON_Origen AS o
            WHERE o.IdEmpresa = @IdEmpresa
              AND (@SoloActivos = 0 OR o.Estado = 1)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR o.CodigoOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR o.NombreOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR o.ModuloOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdOrigen,
            b.CodigoOrigen,
            b.NombreOrigen,
            b.ModuloOrigen,
            b.PermiteRegistroManual,
            b.Estado,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.CodigoOrigen ASC
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
