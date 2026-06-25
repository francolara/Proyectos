-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Lista comprobantes pendientes y notas de credito pendientes de una persona para el modulo Aplicaciones.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Agrega la ayuda central de Aplicaciones para consultar comprobantes con saldo por cliente o proveedor y separar notas de credito segun TipoComprobante 07.

CREATE OR ALTER PROCEDURE dbo.usp_APL_ListarComprobantesPendientesPorPersona
    @IdEmpresa INT,
    @ModuloOperacion VARCHAR(10),
    @IdPersona INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF @ModuloOperacion NOT IN ('COM', 'VEN')
        BEGIN
            RAISERROR(N'El modulo de aplicaciones debe ser COM o VEN.', 16, 1);
        END;

        IF @ModuloOperacion = 'VEN'
        BEGIN
            SELECT
                v.IdVenta AS IdRegistro,
                CAST('VEN' AS VARCHAR(10)) AS ModuloOperacion,
                pe.IdPersona,
                pe.NombreCompleto AS NombrePersona,
                pe.NumeroDocumento AS NumeroDocumentoPersona,
                v.FechaEmision,
                v.TipoComprobante,
                tc.Descripcion AS DescripcionTipoComprobante,
                v.Serie,
                v.Numero,
                v.IdMoneda,
                m.CodigoMoneda,
                v.TipoCambio,
                v.ImporteTotal,
                v.Saldo,
                CAST(CASE WHEN v.TipoComprobante = '07' THEN 1 ELSE 0 END AS BIT) AS EsNotaCredito,
                cfg.EscenarioOperacion,
                ISNULL(v.Observacion, N'') AS Observacion
            FROM dbo.VEN_Venta AS v
            INNER JOIN dbo.ADM_Cliente AS c
                ON c.IdCliente = v.IdCliente
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = c.IdPersona
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = v.IdMoneda
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = v.TipoComprobante
            INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
                ON cfg.IdConfiguracionContabilizacion = v.IdConfiguracionContabilizacion
            WHERE v.IdEmpresa = @IdEmpresa
              AND pe.IdPersona = @IdPersona
              AND v.Saldo > 0
            ORDER BY
                CASE WHEN v.TipoComprobante = '07' THEN 1 ELSE 0 END,
                v.FechaEmision ASC,
                v.IdVenta ASC;
        END
        ELSE
        BEGIN
            SELECT
                c.IdCompra AS IdRegistro,
                CAST('COM' AS VARCHAR(10)) AS ModuloOperacion,
                pe.IdPersona,
                pe.NombreCompleto AS NombrePersona,
                pe.NumeroDocumento AS NumeroDocumentoPersona,
                c.FechaEmision,
                c.TipoComprobante,
                tc.Descripcion AS DescripcionTipoComprobante,
                c.Serie,
                c.Numero,
                c.IdMoneda,
                m.CodigoMoneda,
                c.TipoCambio,
                c.ImporteTotal,
                c.Saldo,
                CAST(CASE WHEN c.TipoComprobante = '07' THEN 1 ELSE 0 END AS BIT) AS EsNotaCredito,
                cfg.EscenarioOperacion,
                ISNULL(c.Observacion, N'') AS Observacion
            FROM dbo.COM_Compra AS c
            INNER JOIN dbo.ADM_Proveedor AS p
                ON p.IdProveedor = c.IdProveedor
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = p.IdPersona
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = c.IdMoneda
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = c.TipoComprobante
            INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
                ON cfg.IdConfiguracionContabilizacion = c.IdConfiguracionContabilizacion
            WHERE c.IdEmpresa = @IdEmpresa
              AND pe.IdPersona = @IdPersona
              AND c.Saldo > 0
            ORDER BY
                CASE WHEN c.TipoComprobante = '07' THEN 1 ELSE 0 END,
                c.FechaEmision ASC,
                c.IdCompra ASC;
        END;

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
