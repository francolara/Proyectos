-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/07/2026
-- Description:   Genera el registro de compras HTML tomando la provision COM_Compra para replicar el formato A4 legacy sin depender de CON_AsientoDetalle.
-- =============================================
-- Firma: FRANCO LARA - 07/07/2026 | Crea el procedimiento del registro de compras con filtro por anio, mes y codigo de persona, usando solo la provision actual.

CREATE OR ALTER PROCEDURE dbo.usp_COM_ReporteRegistroCompras
    @IdEmpresa INT,
    @Anio SMALLINT,
    @Mes TINYINT,
    @CodigoPersona VARCHAR(20) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @CodigoPersonaTrabajo VARCHAR(20) = NULLIF(LTRIM(RTRIM(@CodigoPersona)), '')

        SELECT
            c.FechaEmision,
            c.FechaContabilizacion,
            c.TipoComprobante,
            tc.Descripcion AS DescripcionTipoComprobante,
            c.Serie,
            c.Numero,
            p.CodigoProveedor AS CodigoPersona,
            pe.NumeroDocumento AS NumeroDocumentoPersona,
            pe.NombreCompleto AS NombrePersona,
            m.CodigoMoneda,
            c.TipoCambio,
            CASE WHEN c.TipoComprobante = '07' THEN -c.BaseImponible ELSE c.BaseImponible END AS BaseImponibleGravada,
            CASE WHEN c.TipoComprobante = '07' THEN -c.Igv ELSE c.Igv END AS IgvGravado,
            CAST(0 AS DECIMAL(18,2)) AS BaseImponibleGasto,
            CAST(0 AS DECIMAL(18,2)) AS IgvGasto,
            CAST(0 AS DECIMAL(18,2)) AS BaseImponibleSinCredito,
            CAST(0 AS DECIMAL(18,2)) AS IgvSinCredito,
            CASE WHEN c.TipoComprobante = '07' THEN -c.TotalExonerado ELSE c.TotalExonerado END AS TotalExonerado,
            CASE WHEN c.TipoComprobante = '07' THEN -c.TotalInafecto ELSE c.TotalInafecto END AS TotalInafecto,
            CASE WHEN c.TipoComprobante = '07' THEN -c.OtrosTributos ELSE c.OtrosTributos END AS OtrosTributos,
            CASE WHEN c.TipoComprobante = '07' THEN -c.Icbper ELSE c.Icbper END AS Icbper,
            CASE WHEN c.TipoComprobante = '07' THEN -c.Retencion ELSE c.Retencion END AS Retencion,
            CASE WHEN c.TipoComprobante = '07' THEN -c.ImporteDetraccion ELSE c.ImporteDetraccion END AS ImporteDetraccion,
            CASE WHEN c.TipoComprobante = '07' THEN -c.ImportePercepcion ELSE c.ImportePercepcion END AS ImportePercepcion,
            CASE WHEN c.TipoComprobante = '07' THEN -c.ImporteTotal ELSE c.ImporteTotal END AS ImporteTotal,
            c.Estado,
            ISNULL(c.Observacion, N'') AS Observacion
        FROM dbo.COM_Compra AS c
        INNER JOIN dbo.ADM_Proveedor AS p
            ON p.IdProveedor = c.IdProveedor
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = p.IdPersona
        INNER JOIN dbo.ADM_TipoComprobante AS tc
            ON tc.CodigoTipoComprobante = c.TipoComprobante
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = c.IdMoneda
        WHERE c.IdEmpresa = @IdEmpresa
          AND YEAR(c.FechaContabilizacion) = @Anio
          AND MONTH(c.FechaContabilizacion) = @Mes
          AND (
                @CodigoPersonaTrabajo IS NULL
                OR p.CodigoProveedor = @CodigoPersonaTrabajo
              )
        ORDER BY
            c.FechaEmision,
            c.TipoComprobante,
            c.Serie,
            TRY_CONVERT(BIGINT, c.Numero),
            c.Numero,
            c.IdCompra;

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
