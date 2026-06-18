-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista la configuracion contable automatica de compras y ventas con busqueda y paginacion server-side.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarConfiguracionContabilizacionPorEmpresa
    @IdEmpresa INT,
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
                c.IdConfiguracionContabilizacion,
                c.IdEmpresa,
                c.ModuloOperacion,
                c.EscenarioOperacion,
                c.IdOrigen,
                o.CodigoOrigen,
                o.NombreOrigen,
                c.Descripcion,
                c.GeneraAsientoAutomatico,
                c.UsaTipoCambio,
                c.Activo,
                COUNT(d.IdConfiguracionContabilizacionDetalle) AS CantidadComponentes
            FROM dbo.CON_ConfiguracionContabilizacion AS c
            INNER JOIN dbo.CON_Origen AS o
                ON o.IdOrigen = c.IdOrigen
            LEFT JOIN dbo.CON_ConfiguracionContabilizacionDetalle AS d
                ON d.IdConfiguracionContabilizacion = c.IdConfiguracionContabilizacion
               AND d.Activo = 1
            WHERE c.IdEmpresa = @IdEmpresa
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR c.ModuloOperacion LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.EscenarioOperacion LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR o.CodigoOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR o.NombreOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.Descripcion LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
            GROUP BY
                c.IdConfiguracionContabilizacion,
                c.IdEmpresa,
                c.ModuloOperacion,
                c.EscenarioOperacion,
                c.IdOrigen,
                o.CodigoOrigen,
                o.NombreOrigen,
                c.Descripcion,
                c.GeneraAsientoAutomatico,
                c.UsaTipoCambio,
                c.Activo
        )
        SELECT
            b.IdConfiguracionContabilizacion,
            b.IdEmpresa,
            b.ModuloOperacion,
            b.EscenarioOperacion,
            b.IdOrigen,
            b.CodigoOrigen,
            b.NombreOrigen,
            b.Descripcion,
            b.GeneraAsientoAutomatico,
            b.UsaTipoCambio,
            b.Activo,
            b.CantidadComponentes,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY
            b.ModuloOperacion ASC,
            b.EscenarioOperacion ASC
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
