-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Lista transferencias entre cuentas uniendo el movimiento emisor y receptor registrado en BAN_MovimientoBanco.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Expone los IdAsiento emisor y receptor junto con sus numeros para permitir la navegacion directa al asiento desde el listado de transferencias.

CREATE OR ALTER PROCEDURE dbo.usp_BAN_ListarTransferenciasCuentaPorEmpresa
    @IdEmpresa INT,
    @Anio SMALLINT,
    @Mes TINYINT,
    @TextoBusqueda NVARCHAR(200) = NULL,
    @NumeroPagina INT = 1,
    @TamanoPagina INT = 20
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'');
        DECLARE @FechaInicio DATE = DATEFROMPARTS(@Anio, @Mes, 1);
        DECLARE @FechaFin DATE = DATEADD(MONTH, 1, @FechaInicio);
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE 1 END;
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE 20 END;

        ;WITH Operaciones AS
        (
            SELECT
                LTRIM(RTRIM(op.idOpeBancaria)) AS IdOpeBancaria,
                LTRIM(RTRIM(op.Destino)) AS TipoMovimiento,
                MAX(LTRIM(RTRIM(op.Tipo))) AS TipoOperacion
            FROM dbo.operacionesbancarias AS op
            WHERE LTRIM(RTRIM(op.idTipoOpeBancaria)) = 'T'
            GROUP BY
                LTRIM(RTRIM(op.idOpeBancaria)),
                LTRIM(RTRIM(op.Destino))
        ),
        Base AS
        (
            SELECT
                em.IdTransferenciaCuenta,
                em.IdMovimientoBanco AS IdMovimientoBancoEmisor,
                em.IdAsiento AS IdAsientoEmisor,
                em.NumeroMovimiento AS NumeroMovimientoEmisor,
                aem.NumeroAsiento AS NumeroAsientoEmisor,
                em.IdBancoConfiguracionEmpresa AS IdBancoConfiguracionEmpresaEmisor,
                ccem.NroCuentaCorriente AS CuentaCorrienteEmisor,
                ISNULL(monem.CodigoMoneda, '') AS MonedaEmisor,
                ISNULL(opem.TipoOperacion, '') AS OperacionEmisor,
                em.FechaEmision AS FechaEmisionEmisor,
                em.TipoCambio AS TipoCambioEmisor,
                ISNULL(em.NumeroDocumento, '') AS NumeroOperacionEmisor,
                em.ImporteTotal AS ImporteEmisor,
                em.Glosa AS GlosaEmisor,
                rec.IdMovimientoBanco AS IdMovimientoBancoReceptor,
                rec.IdAsiento AS IdAsientoReceptor,
                rec.NumeroMovimiento AS NumeroMovimientoReceptor,
                arec.NumeroAsiento AS NumeroAsientoReceptor,
                rec.IdBancoConfiguracionEmpresa AS IdBancoConfiguracionEmpresaReceptor,
                ccrec.NroCuentaCorriente AS CuentaCorrienteReceptor,
                ISNULL(monrec.CodigoMoneda, '') AS MonedaReceptor,
                ISNULL(oprec.TipoOperacion, '') AS OperacionReceptor,
                rec.FechaEmision AS FechaEmisionReceptor,
                rec.TipoCambio AS TipoCambioReceptor,
                ISNULL(rec.NumeroDocumento, '') AS NumeroOperacionReceptor,
                rec.ImporteTotal AS ImporteReceptor,
                rec.Glosa AS GlosaReceptor
            FROM dbo.BAN_MovimientoBanco AS em
            INNER JOIN dbo.BAN_MovimientoBanco AS rec
                ON rec.IdTransferenciaCuenta = em.IdTransferenciaCuenta
               AND rec.RolTransferencia = 'I'
               AND rec.Activo = 1
               AND rec.IdEmpresa = em.IdEmpresa
            INNER JOIN dbo.CON_BancosConfiguracionEmpresa AS ccem
                ON ccem.IdBancoConfiguracionEmpresa = em.IdBancoConfiguracionEmpresa
            INNER JOIN dbo.CON_BancosConfiguracionEmpresa AS ccrec
                ON ccrec.IdBancoConfiguracionEmpresa = rec.IdBancoConfiguracionEmpresa
            LEFT JOIN dbo.ADM_Moneda AS monem
                ON monem.IdMoneda = ccem.IdMoneda
            LEFT JOIN dbo.ADM_Moneda AS monrec
                ON monrec.IdMoneda = ccrec.IdMoneda
            LEFT JOIN Operaciones AS opem
                ON opem.IdOpeBancaria = em.IdOpeBancaria
               AND opem.TipoMovimiento = em.TipoMovimiento
            LEFT JOIN Operaciones AS oprec
                ON oprec.IdOpeBancaria = rec.IdOpeBancaria
               AND oprec.TipoMovimiento = rec.TipoMovimiento
            LEFT JOIN dbo.CON_Asiento AS aem
                ON aem.IdAsiento = em.IdAsiento
            LEFT JOIN dbo.CON_Asiento AS arec
                ON arec.IdAsiento = rec.IdAsiento
            WHERE em.IdEmpresa = @IdEmpresa
              AND em.Activo = 1
              AND em.RolTransferencia = 'E'
              AND em.IdTransferenciaCuenta IS NOT NULL
              AND em.FechaEmision >= @FechaInicio
              AND em.FechaEmision < @FechaFin
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR ccem.NroCuentaCorriente LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ccrec.NroCuentaCorriente LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR em.NumeroDocumento LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR rec.NumeroDocumento LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR em.Glosa LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR rec.Glosa LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(opem.TipoOperacion, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(oprec.TipoOperacion, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdTransferenciaCuenta,
            b.IdMovimientoBancoEmisor,
            b.IdAsientoEmisor,
            b.NumeroMovimientoEmisor,
            b.NumeroAsientoEmisor,
            b.IdBancoConfiguracionEmpresaEmisor,
            b.CuentaCorrienteEmisor,
            b.MonedaEmisor,
            b.OperacionEmisor,
            b.FechaEmisionEmisor,
            b.TipoCambioEmisor,
            b.NumeroOperacionEmisor,
            b.ImporteEmisor,
            b.GlosaEmisor,
            b.IdMovimientoBancoReceptor,
            b.IdAsientoReceptor,
            b.NumeroMovimientoReceptor,
            b.NumeroAsientoReceptor,
            b.IdBancoConfiguracionEmpresaReceptor,
            b.CuentaCorrienteReceptor,
            b.MonedaReceptor,
            b.OperacionReceptor,
            b.FechaEmisionReceptor,
            b.TipoCambioReceptor,
            b.NumeroOperacionReceptor,
            b.ImporteReceptor,
            b.GlosaReceptor,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.FechaEmisionEmisor DESC, b.NumeroMovimientoEmisor DESC, b.IdMovimientoBancoEmisor DESC
        OFFSET (@NumeroPaginaTrabajo - 1) * @TamanoPaginaTrabajo ROWS
        FETCH NEXT @TamanoPaginaTrabajo ROWS ONLY;

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
