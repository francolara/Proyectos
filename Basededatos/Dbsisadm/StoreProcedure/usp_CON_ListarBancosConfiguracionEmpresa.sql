-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Lista cuentas corrientes bancarias por empresa con banco, titular, moneda y cuenta contable asociada.
-- =============================================
-- Firma: FRANCO LARA - 09/07/2026 | Expone periodo y saldos iniciales Debe/Haber de la cuenta corriente para su mantenimiento y reutilizacion en Caja y Bancos.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarBancosConfiguracionEmpresa
    @IdEmpresa INT,
    @SoloActivos BIT = 0,
    @TextoBusqueda NVARCHAR(200) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                c.IdBancoConfiguracionEmpresa,
                c.IdEmpresa,
                c.IdBanco,
                b.Codigo AS CodigoBanco,
                b.Nombre AS NombreBanco,
                c.NroCuentaCorriente,
                c.Titular,
                c.IdMoneda,
                m.CodigoMoneda,
                m.NombreMoneda,
                c.IdPlanCuenta,
                p.CodigoCuenta,
                p.NombreCuenta,
                c.PeriodoSaldoInicial,
                c.SaldoInicialDebe,
                c.SaldoInicialHaber,
                c.Activo,
                c.FechaRegistro,
                c.UsuarioRegistro
            FROM dbo.CON_BancosConfiguracionEmpresa AS c
            INNER JOIN dbo.CON_Bancos AS b
                ON b.IdBanco = c.IdBanco
            INNER JOIN dbo.CON_PlanCuenta AS p
                ON p.IdPlanCuenta = c.IdPlanCuenta
            LEFT JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = c.IdMoneda
            WHERE c.IdEmpresa = @IdEmpresa
              AND (@SoloActivos = 0 OR c.Activo = 1)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR b.Codigo LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR b.Nombre LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.NroCuentaCorriente LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR c.Titular LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR m.CodigoMoneda LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR m.NombreMoneda LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR p.CodigoCuenta LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR p.NombreCuenta LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdBancoConfiguracionEmpresa,
            b.IdEmpresa,
            b.IdBanco,
            b.CodigoBanco,
            b.NombreBanco,
            b.NroCuentaCorriente,
            b.Titular,
            b.IdMoneda,
            b.CodigoMoneda,
            b.NombreMoneda,
            b.IdPlanCuenta,
            b.CodigoCuenta,
            b.NombreCuenta,
            b.PeriodoSaldoInicial,
            b.SaldoInicialDebe,
            b.SaldoInicialHaber,
            b.Activo,
            b.FechaRegistro,
            b.UsuarioRegistro,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.NombreBanco ASC, b.NroCuentaCorriente ASC
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
