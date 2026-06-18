-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista las ventas por empresa con filtro por periodo, busqueda y paginacion server-side.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_VEN_ListarVentasPorEmpresa
    @IdEmpresa INT,
    @Periodo CHAR(6) = NULL,
    @Ejercicio SMALLINT = NULL,
    @Mes TINYINT = NULL,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @PeriodoTrabajo CHAR(6) =
            CASE
                WHEN @Periodo IS NOT NULL THEN @Periodo
                WHEN @Ejercicio IS NOT NULL AND @Mes IS NOT NULL THEN CONVERT(CHAR(4), @Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), @Mes), 2)
                ELSE NULL
            END
        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                v.IdVenta,
                v.IdEmpresa,
                v.IdCliente,
                c.CodigoCliente,
                pe.NombreCompleto AS NombreCliente,
                v.IdConfiguracionContabilizacion,
                cfg.ModuloOperacion,
                cfg.EscenarioOperacion,
                v.IdAsiento,
                v.FechaEmision,
                v.FechaContabilizacion,
                CONVERT(CHAR(6), YEAR(v.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(v.FechaContabilizacion)), 2) AS Periodo,
                v.TipoComprobante,
                v.Serie,
                v.Numero,
                v.IdMoneda,
                m.CodigoMoneda,
                v.TipoCambio,
                v.BaseImponible,
                v.Igv,
                v.Isc,
                v.OtrosTributos,
                v.Redondeo,
                v.ImporteTotal,
                v.Observacion,
                v.Estado
            FROM dbo.VEN_Venta AS v
            INNER JOIN dbo.ADM_Cliente AS c
                ON c.IdCliente = v.IdCliente
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = c.IdPersona
            INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
                ON cfg.IdConfiguracionContabilizacion = v.IdConfiguracionContabilizacion
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = v.IdMoneda
            WHERE v.IdEmpresa = @IdEmpresa
              AND (
                    @PeriodoTrabajo IS NULL
                    OR CONVERT(CHAR(6), YEAR(v.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(v.FechaContabilizacion)), 2) = @PeriodoTrabajo
                  )
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR c.CodigoCliente LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR pe.NombreCompleto LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR cfg.EscenarioOperacion LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR v.TipoComprobante LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR v.Serie LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR v.Numero LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(v.Observacion, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdVenta,
            b.IdEmpresa,
            b.IdCliente,
            b.CodigoCliente,
            b.NombreCliente,
            b.IdConfiguracionContabilizacion,
            b.ModuloOperacion,
            b.EscenarioOperacion,
            b.IdAsiento,
            b.FechaEmision,
            b.FechaContabilizacion,
            b.Periodo,
            b.TipoComprobante,
            b.Serie,
            b.Numero,
            b.IdMoneda,
            b.CodigoMoneda,
            b.TipoCambio,
            b.BaseImponible,
            b.Igv,
            b.Isc,
            b.OtrosTributos,
            b.Redondeo,
            b.ImporteTotal,
            b.Observacion,
            b.Estado,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY
            b.FechaContabilizacion DESC,
            b.IdVenta DESC
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
