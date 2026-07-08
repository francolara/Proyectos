-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/07/2026
-- Description:   Genera el registro de ventas HTML tomando la provision VEN_Venta para replicar el formato A4 legacy sin depender de CON_AsientoDetalle.
-- =============================================
-- Firma: FRANCO LARA - 07/07/2026 | Crea el procedimiento del registro de ventas con filtro por anio, mes y codigo de persona, usando solo la provision actual.

CREATE OR ALTER PROCEDURE dbo.usp_VEN_ReporteRegistroVentas
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
            v.FechaEmision,
            v.FechaContabilizacion,
            v.TipoComprobante,
            tc.Descripcion AS DescripcionTipoComprobante,
            v.Serie,
            v.Numero,
            c.CodigoCliente AS CodigoPersona,
            pe.NumeroDocumento AS NumeroDocumentoPersona,
            pe.NombreCompleto AS NombrePersona,
            m.CodigoMoneda,
            v.TipoCambio,
            CASE WHEN v.TipoComprobante = '07' THEN -v.BaseImponible ELSE v.BaseImponible END AS BaseImponible,
            CAST(0 AS DECIMAL(18,2)) AS Descuento,
            CASE WHEN v.TipoComprobante = '07' THEN -v.TotalExonerado ELSE v.TotalExonerado END AS TotalExonerado,
            CASE WHEN v.TipoComprobante = '07' THEN -v.TotalInafecto ELSE v.TotalInafecto END AS TotalInafecto,
            CASE WHEN v.TipoComprobante = '07' THEN -v.Igv ELSE v.Igv END AS Igv,
            CASE WHEN v.TipoComprobante = '07' THEN -v.Isc ELSE v.Isc END AS Isc,
            CASE WHEN v.TipoComprobante = '07' THEN -v.OtrosTributos ELSE v.OtrosTributos END AS OtrosTributos,
            CASE WHEN v.TipoComprobante = '07' THEN -v.Icbper ELSE v.Icbper END AS Icbper,
            CASE WHEN v.TipoComprobante = '07' THEN -v.Redondeo ELSE v.Redondeo END AS Redondeo,
            CASE WHEN v.TipoComprobante = '07' THEN -v.ImporteTotal ELSE v.ImporteTotal END AS ImporteTotal,
            v.Estado,
            ISNULL(v.Observacion, N'') AS Observacion
        FROM dbo.VEN_Venta AS v
        INNER JOIN dbo.ADM_Cliente AS c
            ON c.IdCliente = v.IdCliente
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = c.IdPersona
        INNER JOIN dbo.ADM_TipoComprobante AS tc
            ON tc.CodigoTipoComprobante = v.TipoComprobante
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = v.IdMoneda
        WHERE v.IdEmpresa = @IdEmpresa
          AND YEAR(v.FechaContabilizacion) = @Anio
          AND MONTH(v.FechaContabilizacion) = @Mes
          AND (
                @CodigoPersonaTrabajo IS NULL
                OR c.CodigoCliente = @CodigoPersonaTrabajo
              )
        ORDER BY
            v.FechaEmision,
            v.TipoComprobante,
            v.Serie,
            TRY_CONVERT(BIGINT, v.Numero),
            v.Numero,
            v.IdVenta;

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
