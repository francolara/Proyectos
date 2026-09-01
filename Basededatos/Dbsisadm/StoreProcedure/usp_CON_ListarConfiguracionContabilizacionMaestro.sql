-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Lista configuraciones contables maestras y el origen asignado para su mantenimiento.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarConfiguracionContabilizacionMaestro
    @TextoBusqueda NVARCHAR(200) = NULL,
    @NumeroPagina INT = 1,
    @TamanoPagina INT = 20
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Pagina INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE 1 END;
        DECLARE @Tamano INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE 20 END;
        DECLARE @Filtro NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'');

        ;WITH Base AS
        (
            SELECT
                configuracion.IdConfiguracionContabilizacionMaestro,
                configuracion.ModuloOperacion,
                configuracion.EscenarioOperacion,
                configuracion.CodigoOrigen,
                origen.NombreOrigen,
                configuracion.Descripcion,
                configuracion.GeneraAsientoAutomatico,
                configuracion.UsaTipoCambio,
                configuracion.Activo,
                configuracion.Orden
            FROM dbo.CON_ConfiguracionContabilizacionMaestro AS configuracion
            LEFT JOIN dbo.CON_OrigenMaestro AS origen
                ON origen.CodigoOrigen = configuracion.CodigoOrigen
            WHERE
                @Filtro IS NULL
                OR configuracion.ModuloOperacion LIKE '%' + @Filtro + '%'
                OR configuracion.EscenarioOperacion LIKE '%' + @Filtro + '%'
                OR configuracion.Descripcion LIKE '%' + @Filtro + '%'
                OR configuracion.CodigoOrigen LIKE '%' + @Filtro + '%'
                OR origen.NombreOrigen LIKE '%' + @Filtro + '%'
        )
        SELECT
            base.IdConfiguracionContabilizacionMaestro,
            base.ModuloOperacion,
            base.EscenarioOperacion,
            base.CodigoOrigen,
            base.NombreOrigen,
            base.Descripcion,
            base.GeneraAsientoAutomatico,
            base.UsaTipoCambio,
            base.Activo,
            base.Orden,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS base
        ORDER BY base.Orden, base.ModuloOperacion, base.EscenarioOperacion
        OFFSET (@Pagina - 1) * @Tamano ROWS
        FETCH NEXT @Tamano ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
