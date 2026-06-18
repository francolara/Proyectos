-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista las reglas de cuentas de destino por empresa y ejercicio con resumen de porcentajes.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Agrega busqueda y paginacion server-side para el mantenimiento web de reglas por ejercicio.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarCuentasDestinoReglaPorEmpresa
    @IdEmpresa INT,
    @Ejercicio SMALLINT,
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
                r.IdCuentaDestinoRegla,
                r.IdEmpresa,
                r.Ejercicio,
                r.IdPlanCuentaOrigen,
                po.CodigoCuenta AS CodigoCuentaOrigen,
                po.NombreCuenta AS NombreCuentaOrigen,
                r.Activo,
                r.Observacion,
                COUNT(d.IdCuentaDestinoReglaDetalle) AS CantidadTramos,
                COALESCE(SUM(d.Porcentaje), 0) AS PorcentajeTotal
            FROM dbo.CON_CuentaDestinoRegla AS r
            INNER JOIN dbo.CON_PlanCuenta AS po
                ON po.IdPlanCuenta = r.IdPlanCuentaOrigen
            LEFT JOIN dbo.CON_CuentaDestinoReglaDetalle AS d
                ON d.IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
               AND d.Activo = 1
            WHERE r.IdEmpresa = @IdEmpresa
              AND r.Ejercicio = @Ejercicio
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR po.CodigoCuenta LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR po.NombreCuenta LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(r.Observacion, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
            GROUP BY
                r.IdCuentaDestinoRegla,
                r.IdEmpresa,
                r.Ejercicio,
                r.IdPlanCuentaOrigen,
                po.CodigoCuenta,
                po.NombreCuenta,
                r.Activo,
                r.Observacion
        )
        SELECT
            b.IdCuentaDestinoRegla,
            b.IdEmpresa,
            b.Ejercicio,
            b.IdPlanCuentaOrigen,
            b.CodigoCuentaOrigen,
            b.NombreCuentaOrigen,
            b.Activo,
            b.Observacion,
            b.CantidadTramos,
            b.PorcentajeTotal,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.CodigoCuentaOrigen ASC
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
