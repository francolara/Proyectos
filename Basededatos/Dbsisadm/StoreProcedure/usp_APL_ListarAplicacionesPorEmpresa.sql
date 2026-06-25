-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Lista aplicaciones de notas de credito por empresa, periodo y filtros operativos.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Agrega el listado paginado del modulo Aplicaciones mostrando comprobante, nota de credito, persona, importe aplicado y asiento generado.

CREATE OR ALTER PROCEDURE dbo.usp_APL_ListarAplicacionesPorEmpresa
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
            END;
        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '');
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END;
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END;

        ;WITH Base AS
        (
            SELECT
                a.IdAplicacionNotaCredito,
                a.IdEmpresa,
                a.ModuloOperacion,
                a.IdPersona,
                p.NombreCompleto AS NombrePersona,
                p.NumeroDocumento AS NumeroDocumentoPersona,
                a.FechaAplicacion,
                CONVERT(CHAR(4), YEAR(a.FechaAplicacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(a.FechaAplicacion)), 2) AS Periodo,
                a.IdRegistroComprobante,
                a.IdRegistroNotaCredito,
                a.IdMoneda,
                m.CodigoMoneda,
                a.TipoCambio,
                a.ImporteAplicado,
                a.IdAsiento,
                asi.NumeroAsiento,
                a.Glosa,
                a.Observacion,
                CASE
                    WHEN a.ModuloOperacion = 'VEN' THEN N'Cliente'
                    ELSE N'Proveedor'
                END AS TipoPersonaTexto,
                COALESCE(vc.TipoComprobante, cc.TipoComprobante) AS TipoComprobanteAplicado,
                COALESCE(tca.Descripcion, tcb.Descripcion) AS DescripcionTipoComprobanteAplicado,
                COALESCE(vc.Serie, cc.Serie) AS SerieAplicado,
                COALESCE(vc.Numero, cc.Numero) AS NumeroAplicado,
                COALESCE(vn.TipoComprobante, cn.TipoComprobante) AS TipoComprobanteNc,
                COALESCE(tcnv.Descripcion, tcnc.Descripcion) AS DescripcionTipoComprobanteNc,
                COALESCE(vn.Serie, cn.Serie) AS SerieNc,
                COALESCE(vn.Numero, cn.Numero) AS NumeroNc,
                COUNT(1) OVER() AS TotalRegistros
            FROM dbo.CON_AplicacionNotaCredito AS a
            INNER JOIN dbo.ADM_Persona AS p
                ON p.IdPersona = a.IdPersona
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = a.IdMoneda
            LEFT JOIN dbo.CON_Asiento AS asi
                ON asi.IdAsiento = a.IdAsiento
            LEFT JOIN dbo.VEN_Venta AS vc
                ON a.ModuloOperacion = 'VEN'
               AND vc.IdVenta = a.IdRegistroComprobante
            LEFT JOIN dbo.VEN_Venta AS vn
                ON a.ModuloOperacion = 'VEN'
               AND vn.IdVenta = a.IdRegistroNotaCredito
            LEFT JOIN dbo.COM_Compra AS cc
                ON a.ModuloOperacion = 'COM'
               AND cc.IdCompra = a.IdRegistroComprobante
            LEFT JOIN dbo.COM_Compra AS cn
                ON a.ModuloOperacion = 'COM'
               AND cn.IdCompra = a.IdRegistroNotaCredito
            LEFT JOIN dbo.ADM_TipoComprobante AS tca
                ON tca.CodigoTipoComprobante = vc.TipoComprobante
            LEFT JOIN dbo.ADM_TipoComprobante AS tcb
                ON tcb.CodigoTipoComprobante = cc.TipoComprobante
            LEFT JOIN dbo.ADM_TipoComprobante AS tcnv
                ON tcnv.CodigoTipoComprobante = vn.TipoComprobante
            LEFT JOIN dbo.ADM_TipoComprobante AS tcnc
                ON tcnc.CodigoTipoComprobante = cn.TipoComprobante
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.Activo = 1
              AND (
                    @PeriodoTrabajo IS NULL
                    OR CONVERT(CHAR(4), YEAR(a.FechaAplicacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(a.FechaAplicacion)), 2) = @PeriodoTrabajo
                  )
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR p.NombreCompleto LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR p.NumeroDocumento LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR a.Glosa LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR COALESCE(vc.Serie, cc.Serie, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR COALESCE(vc.Numero, cc.Numero, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR COALESCE(vn.Serie, cn.Serie, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR COALESCE(vn.Numero, cn.Numero, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdAplicacionNotaCredito,
            b.IdEmpresa,
            b.ModuloOperacion,
            b.IdPersona,
            b.NombrePersona,
            b.NumeroDocumentoPersona,
            b.TipoPersonaTexto,
            b.FechaAplicacion,
            b.Periodo,
            b.IdRegistroComprobante,
            b.IdRegistroNotaCredito,
            b.IdMoneda,
            b.CodigoMoneda,
            b.TipoCambio,
            b.ImporteAplicado,
            b.IdAsiento,
            b.NumeroAsiento,
            b.Glosa,
            b.Observacion,
            b.TipoComprobanteAplicado,
            b.DescripcionTipoComprobanteAplicado,
            b.SerieAplicado,
            b.NumeroAplicado,
            b.TipoComprobanteNc,
            b.DescripcionTipoComprobanteNc,
            b.SerieNc,
            b.NumeroNc,
            b.TotalRegistros
        FROM Base AS b
        ORDER BY
            b.FechaAplicacion DESC,
            b.IdAplicacionNotaCredito DESC
        OFFSET CASE WHEN @NumeroPaginaTrabajo IS NULL OR @TamanoPaginaTrabajo IS NULL THEN 0 ELSE (@NumeroPaginaTrabajo - 1) * @TamanoPaginaTrabajo END ROWS
        FETCH NEXT CASE WHEN @NumeroPaginaTrabajo IS NULL OR @TamanoPaginaTrabajo IS NULL THEN 2147483647 ELSE @TamanoPaginaTrabajo END ROWS ONLY;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);

    END CATCH

END
