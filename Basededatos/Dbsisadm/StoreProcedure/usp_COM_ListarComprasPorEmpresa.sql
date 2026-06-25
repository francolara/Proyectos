-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista las provisiones de compra por empresa con filtro por periodo, busqueda y paginacion server-side.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Incluye subtotal, totales exonerado/inafecto e IGV en el listado de provisiones de compra.
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
-- Description:   Agrega la situacion del comprobante de compra segun el saldo pendiente.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_COM_ListarComprasPorEmpresa
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
                c.IdCompra,
                c.IdEmpresa,
                c.IdProveedor,
                p.CodigoProveedor,
                pe.NombreCompleto AS NombreProveedor,
                c.IdConfiguracionContabilizacion,
                cfg.ModuloOperacion,
                cfg.EscenarioOperacion,
                c.IdAsiento,
                c.FechaEmision,
                c.FechaContabilizacion,
                CONVERT(CHAR(4), YEAR(c.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(c.FechaContabilizacion)), 2) AS Periodo,
                c.TipoComprobante,
                tc.Descripcion AS DescripcionTipoComprobante,
                c.Serie,
                c.Numero,
                pe.NumeroDocumento AS NumeroDocumentoPersona,
                c.IdMoneda,
                m.CodigoMoneda,
                c.TipoCambio,
                c.BaseImponible,
                c.TotalExonerado,
                c.TotalInafecto,
                c.Icbper,
                c.Igv,
                c.Isc,
                c.OtrosTributos,
                c.Redondeo,
                c.ImporteTotal,
                c.Saldo,
                c.Observacion,
                c.Estado,
                CASE
                    WHEN c.Saldo <= 0 THEN N'Pagada'
                    WHEN c.Saldo < c.ImporteTotal THEN N'Pagada Parcial'
                    ELSE N'Pendiente'
                END AS Situacion
            FROM dbo.COM_Compra AS c
            INNER JOIN dbo.ADM_Proveedor AS p
                ON p.IdProveedor = c.IdProveedor
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = p.IdPersona
            INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
                ON cfg.IdConfiguracionContabilizacion = c.IdConfiguracionContabilizacion
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = c.IdMoneda
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = c.TipoComprobante
            WHERE c.IdEmpresa = @IdEmpresa
              AND (
                    @PeriodoTrabajo IS NULL
                    OR CONVERT(CHAR(4), YEAR(c.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(c.FechaContabilizacion)), 2) = @PeriodoTrabajo
                  )
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR p.CodigoProveedor LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR pe.NombreCompleto LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR cfg.EscenarioOperacion LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.TipoComprobante LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.Serie LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.Numero LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(c.Observacion, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdCompra,
            b.IdEmpresa,
            b.IdProveedor,
            b.CodigoProveedor,
            b.NombreProveedor,
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
            b.IdCompra DESC
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
