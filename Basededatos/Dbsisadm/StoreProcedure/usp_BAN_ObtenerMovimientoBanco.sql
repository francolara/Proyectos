-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Obtiene cabecera y detalle de un movimiento de caja y bancos usando NombreCompleto de ADM_Persona, correlativo interno mensual y referencias documentarias por linea.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Amplia la obtencion de Caja y Bancos devolviendo TipoCambio y Observacion de la cabecera, mas persona por linea y origen/aplicacion de comprobantes en el detalle.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Expone TotalImporteS y TotalImporteD del detalle para mantener consistencia con el nuevo modelo por moneda.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Incluye el asiento contable vinculado para consulta y edicion del movimiento bancario.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_BAN_ObtenerMovimientoBanco
    @IdMovimientoBanco INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            m.IdMovimientoBanco,
            m.IdAsiento,
            m.IdEmpresa,
            m.IdBancoConfiguracionEmpresa,
            m.NumeroMovimiento,
            cc.NroCuentaCorriente,
            b.Codigo AS CodigoBanco,
            b.Nombre AS NombreBanco,
            cc.Titular AS TitularCuentaCorriente,
            cc.IdMoneda,
            mon.CodigoMoneda,
            mon.NombreMoneda,
            m.TipoMovimiento,
            m.IdOpeBancaria,
            m.FechaEmision,
            m.TipoCambio,
            m.IdPersona,
            p.NumeroDocumento AS NumeroDocumentoPersona,
            CASE
                WHEN p.TipoPersona = 'J'
                    THEN p.RazonSocial
                ELSE NULLIF(LTRIM(RTRIM(ISNULL(p.NombreCompleto, ''))), '')
            END AS NombrePersona,
            m.NumeroDocumento,
            m.Glosa,
            m.Observacion,
            m.ImporteTotal,
            a.NumeroAsiento,
            m.Activo
        FROM dbo.BAN_MovimientoBanco AS m
        INNER JOIN dbo.CON_BancosConfiguracionEmpresa AS cc
            ON cc.IdBancoConfiguracionEmpresa = m.IdBancoConfiguracionEmpresa
        INNER JOIN dbo.CON_Bancos AS b
            ON b.IdBanco = cc.IdBanco
        LEFT JOIN dbo.ADM_Moneda AS mon
            ON mon.IdMoneda = cc.IdMoneda
        LEFT JOIN dbo.ADM_Persona AS p
            ON p.IdPersona = m.IdPersona
        LEFT JOIN dbo.CON_Asiento AS a
            ON a.IdAsiento = m.IdAsiento
        WHERE m.IdMovimientoBanco = @IdMovimientoBanco
          AND m.IdEmpresa = @IdEmpresa
          AND m.Activo = 1;

        SELECT
            d.IdMovimientoBancoDetalle,
            d.IdMovimientoBanco,
            d.Item,
            d.IdPlanCuenta,
            d.IdPersona,
            d.ModuloOperacionComprobante,
            d.IdRegistroComprobante,
            d.ImporteAplicado,
            pd.NumeroDocumento AS NumeroDocumentoPersona,
            CASE
                WHEN pd.TipoPersona = 'J'
                    THEN pd.RazonSocial
                ELSE NULLIF(LTRIM(RTRIM(ISNULL(pd.NombreCompleto, ''))), '')
            END AS NombrePersona,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            pc.RequiereCentroCosto,
            d.GlosaDetalle,
            d.CodigoCentroCosto,
            d.NumeroDocumento,
            d.TipoDocumento,
            d.Serie,
            d.ReferenciaLinea,
            d.TipoCambioLinea,
            d.Debe,
            d.Haber,
            d.TotalImporteS,
            d.TotalImporteD
        FROM dbo.BAN_MovimientoBancoDetalle AS d
        INNER JOIN dbo.BAN_MovimientoBanco AS m
            ON m.IdMovimientoBanco = d.IdMovimientoBanco
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = d.IdPlanCuenta
        LEFT JOIN dbo.ADM_Persona AS pd
            ON pd.IdPersona = d.IdPersona
        WHERE d.IdMovimientoBanco = @IdMovimientoBanco
          AND m.IdEmpresa = @IdEmpresa
        ORDER BY d.Item ASC;

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
