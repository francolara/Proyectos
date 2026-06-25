-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista las ventas por empresa con filtro por periodo, busqueda y paginacion server-side.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Incorpora subtotal, total exonerado, total inafecto e ICBPER interno en el listado de ventas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Corrige el armado del periodo yyyyMM para evitar espacios y permitir el filtro correcto por mes.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Devuelve saldo, descripcion del comprobante y numero de documento de la persona para ayudas y control de comprobantes.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Unifica el estado provisionado de ventas y agrega la situacion del comprobante segun el saldo pendiente.
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

        DECLARE @PeriodoTrabajo VARCHAR(6) =
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
                CONVERT(CHAR(4), YEAR(v.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(v.FechaContabilizacion)), 2) AS Periodo,
                v.TipoComprobante,
                tc.Descripcion AS DescripcionTipoComprobante,
                v.Serie,
                v.Numero,
                pe.NumeroDocumento AS NumeroDocumentoPersona,
                v.IdMoneda,
                m.CodigoMoneda,
                v.TipoCambio,
                v.BaseImponible,
                v.TotalExonerado,
                v.TotalInafecto,
                v.Icbper,
                v.Igv,
                v.Isc,
                v.OtrosTributos,
                v.Redondeo,
                v.ImporteTotal,
                v.Saldo,
                v.Observacion,
                v.Estado,
                CASE
                    WHEN v.Saldo <= 0 THEN N'Pagada'
                    WHEN v.Saldo < v.ImporteTotal THEN N'Pagada Parcial'
                    ELSE N'Pendiente'
                END AS Situacion
            FROM dbo.VEN_Venta AS v
            INNER JOIN dbo.ADM_Cliente AS c
                ON c.IdCliente = v.IdCliente
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = c.IdPersona
            INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
                ON cfg.IdConfiguracionContabilizacion = v.IdConfiguracionContabilizacion
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = v.IdMoneda
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = v.TipoComprobante
            WHERE v.IdEmpresa = @IdEmpresa
              AND (
                    @PeriodoTrabajo IS NULL
                    OR CONVERT(CHAR(4), YEAR(v.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(v.FechaContabilizacion)), 2) = @PeriodoTrabajo
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
            b.DescripcionTipoComprobante,
            b.Serie,
            b.Numero,
            b.NumeroDocumentoPersona,
            b.IdMoneda,
            b.CodigoMoneda,
            b.TipoCambio,
            b.BaseImponible,
            b.TotalExonerado,
            b.TotalInafecto,
            b.Icbper,
            b.Igv,
            b.Isc,
            b.OtrosTributos,
            b.Redondeo,
            b.ImporteTotal,
            b.Saldo,
            b.Observacion,
            b.Estado,
            b.Situacion,
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
