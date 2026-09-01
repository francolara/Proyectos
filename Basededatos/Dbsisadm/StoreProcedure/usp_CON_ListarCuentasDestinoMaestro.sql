-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Lista las reglas maestras de cuentas destino con resumen y paginacion.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarCuentasDestinoMaestro
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
                regla.IdCuentaDestinoReglaMaestro,
                regla.CodigoCuentaOrigen,
                cuenta.NombreCuenta AS NombreCuentaOrigen,
                regla.Activo,
                regla.Observacion,
                COUNT(detalle.IdCuentaDestinoReglaDetalleMaestro) AS CantidadTramos,
                ISNULL(SUM(CASE WHEN detalle.Activo = 1 THEN detalle.Porcentaje ELSE 0 END), 0) AS PorcentajeTotal
            FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
            LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
                ON cuenta.CodigoCuenta = regla.CodigoCuentaOrigen
            LEFT JOIN dbo.CON_CuentaDestinoReglaDetalleMaestro AS detalle
                ON detalle.IdCuentaDestinoReglaMaestro = regla.IdCuentaDestinoReglaMaestro
            WHERE @Filtro IS NULL
               OR regla.CodigoCuentaOrigen LIKE '%' + @Filtro + '%'
               OR cuenta.NombreCuenta LIKE '%' + @Filtro + '%'
               OR regla.Observacion LIKE '%' + @Filtro + '%'
            GROUP BY regla.IdCuentaDestinoReglaMaestro, regla.CodigoCuentaOrigen, cuenta.NombreCuenta, regla.Activo, regla.Observacion
        )
        SELECT
            base.IdCuentaDestinoReglaMaestro,
            base.CodigoCuentaOrigen,
            base.NombreCuentaOrigen,
            base.Activo,
            base.Observacion,
            base.CantidadTramos,
            base.PorcentajeTotal,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS base
        ORDER BY base.CodigoCuentaOrigen
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
