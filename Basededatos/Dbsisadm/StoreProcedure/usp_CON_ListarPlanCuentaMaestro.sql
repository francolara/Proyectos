-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Lista el plan de cuentas maestro con filtros y paginacion para el panel SuperAdmin.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarPlanCuentaMaestro
    @IdPlanCuentaMaestro INT = NULL,
    @TextoBusqueda NVARCHAR(200) = NULL,
    @NivelCuenta TINYINT = NULL,
    @SoloMovimiento BIT = 0,
    @SoloActivos BIT = 0,
    @NumeroPagina INT = 1,
    @TamanoPagina INT = 20
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Pagina INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE 1 END;
        DECLARE @Tamano INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE 20 END;
        DECLARE @Filtro NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'');

        ;WITH Base AS
        (
            SELECT
                cuenta.IdPlanCuentaMaestro,
                cuenta.CodigoCuenta,
                cuenta.CodigoCuentaPadre,
                cuenta.NombreCuenta,
                cuenta.NivelCuenta,
                cuenta.ColBalance,
                cuenta.IdMoneda,
                cuenta.TipoCambio,
                cuenta.AceptaMovimiento,
                cuenta.RequiereCentroCosto,
                cuenta.Estado,
                cuenta.Orden,
                CAST(CASE WHEN EXISTS
                (
                    SELECT 1
                    FROM dbo.CON_PlanCuentaMaestro AS hija
                    WHERE hija.CodigoCuentaPadre = cuenta.CodigoCuenta
                ) THEN 0 ELSE 1 END AS BIT) AS EsUltimoNivel
            FROM dbo.CON_PlanCuentaMaestro AS cuenta
            WHERE (@IdPlanCuentaMaestro IS NULL OR cuenta.IdPlanCuentaMaestro = @IdPlanCuentaMaestro)
              AND (@NivelCuenta IS NULL OR cuenta.NivelCuenta = @NivelCuenta)
              AND (@SoloMovimiento = 0 OR cuenta.AceptaMovimiento = 1)
              AND (@SoloActivos = 0 OR cuenta.Estado = 1)
              AND
              (
                  @Filtro IS NULL
                  OR cuenta.CodigoCuenta LIKE '%' + @Filtro + '%'
                  OR cuenta.NombreCuenta LIKE '%' + @Filtro + '%'
              )
        )
        SELECT
            base.IdPlanCuentaMaestro,
            base.CodigoCuenta,
            base.CodigoCuentaPadre,
            base.NombreCuenta,
            base.NivelCuenta,
            base.ColBalance,
            base.IdMoneda,
            base.TipoCambio,
            base.AceptaMovimiento,
            base.RequiereCentroCosto,
            base.Estado,
            base.Orden,
            base.EsUltimoNivel,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS base
        ORDER BY base.Orden, base.CodigoCuenta
        OFFSET (@Pagina - 1) * @Tamano ROWS
        FETCH NEXT @Tamano ROWS ONLY;
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
