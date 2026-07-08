-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Lista movimientos de caja y bancos por empresa, periodo y cuenta corriente incluyendo correlativo interno mensual.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega el numero de asiento contable vinculado al listado de movimientos bancarios.
-- =============================================
-- Firma: FRANCO LARA - 03/07/2026 | Cambia el filtro del listado para usar el Periodo persistido del movimiento bancario en lugar del rango directo por FechaEmision.

CREATE OR ALTER PROCEDURE dbo.usp_BAN_ListarMovimientosBancoPorEmpresa
    @IdEmpresa INT,
    @IdBancoConfiguracionEmpresa INT = NULL,
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
        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), @Anio) + RIGHT('0' + CONVERT(VARCHAR(2), @Mes), 2);
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE 1 END;
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE 20 END;

        ;WITH Operaciones AS
        (
            SELECT
                LTRIM(RTRIM(op.idOpeBancaria)) AS IdOpeBancaria,
                LTRIM(RTRIM(op.Destino)) AS TipoMovimiento,
                MAX(LTRIM(RTRIM(op.Tipo))) AS TipoOperacion
            FROM dbo.operacionesbancarias AS op
            GROUP BY
                LTRIM(RTRIM(op.idOpeBancaria)),
                LTRIM(RTRIM(op.Destino))
        ),
        Base AS
        (
            SELECT
                m.IdMovimientoBanco,
                m.IdAsiento,
                m.IdEmpresa,
                m.IdBancoConfiguracionEmpresa,
                m.NumeroMovimiento,
                cc.NroCuentaCorriente,
                cc.Titular,
                b.IdBanco,
                b.Codigo AS CodigoBanco,
                b.Nombre AS NombreBanco,
                mon.CodigoMoneda,
                mon.NombreMoneda,
                m.TipoMovimiento,
                m.IdOpeBancaria,
                ISNULL(op.TipoOperacion, '') AS TipoOperacion,
                m.FechaEmision,
                m.IdPersona,
                p.NumeroDocumento AS NumeroDocumentoPersona,
                CASE
                    WHEN p.TipoPersona = 'J'
                        THEN p.RazonSocial
                    ELSE NULLIF(LTRIM(RTRIM(ISNULL(p.NombreCompleto, ''))), '')
                END AS NombrePersona,
                m.NumeroDocumento,
                m.Glosa,
                m.ImporteTotal,
                a.NumeroAsiento,
                CAST(CASE WHEN m.TipoMovimiento = 'I' THEN m.ImporteTotal ELSE 0 END AS DECIMAL(18, 2)) AS Ingreso,
                CAST(CASE WHEN m.TipoMovimiento = 'E' THEN m.ImporteTotal ELSE 0 END AS DECIMAL(18, 2)) AS Egreso,
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
            LEFT JOIN Operaciones AS op
                ON op.IdOpeBancaria = m.IdOpeBancaria
               AND op.TipoMovimiento = m.TipoMovimiento
            LEFT JOIN dbo.CON_Asiento AS a
                ON a.IdAsiento = m.IdAsiento
            WHERE m.IdEmpresa = @IdEmpresa
              AND m.Activo = 1
              AND (@IdBancoConfiguracionEmpresa IS NULL OR m.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa)
              AND m.Periodo = @Periodo
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR cc.NroCuentaCorriente LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR b.Codigo LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR b.Nombre LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR m.NumeroDocumento LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR m.Glosa LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR p.NumeroDocumento LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(op.TipoOperacion, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdMovimientoBanco,
            b.IdAsiento,
            b.IdEmpresa,
            b.IdBancoConfiguracionEmpresa,
            b.NumeroMovimiento,
            b.NroCuentaCorriente,
            b.Titular,
            b.IdBanco,
            b.CodigoBanco,
            b.NombreBanco,
            b.CodigoMoneda,
            b.NombreMoneda,
            b.TipoMovimiento,
            b.IdOpeBancaria,
            b.TipoOperacion,
            b.FechaEmision,
            b.IdPersona,
            b.NumeroDocumentoPersona,
            b.NombrePersona,
            b.NumeroDocumento,
            b.Glosa,
            b.ImporteTotal,
            b.NumeroAsiento,
            b.Ingreso,
            b.Egreso,
            b.Activo,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.FechaEmision DESC, b.NumeroMovimiento DESC, b.IdMovimientoBanco DESC
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
