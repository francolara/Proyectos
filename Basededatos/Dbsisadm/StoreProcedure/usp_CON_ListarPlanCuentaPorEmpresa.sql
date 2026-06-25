-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista el plan de cuentas de una empresa con filtro opcional para cuentas de movimiento.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Ajusta la salida para incluir estado, busqueda y paginacion server-side en el mantenimiento web.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Reemplaza NaturalezaSaldo por ColBalance, agrega IdMoneda/TipoCambio y filtro por nivel.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega indicador de ultimo nivel y permite filtrar ayudas para seleccionar solo cuentas hoja.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarPlanCuentaPorEmpresa
    @IdEmpresa INT,
    @SoloMovimiento BIT = 0,
    @SoloUltimoNivel BIT = 0,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @NivelCuenta TINYINT = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                pc.IdPlanCuenta,
                pc.IdPlanCuentaPadre,
                pc.CodigoCuenta,
                pc.NombreCuenta,
                pc.NivelCuenta,
                pc.ColBalance,
                pc.IdMoneda,
                pc.TipoCambio,
                pc.AceptaMovimiento,
                CAST(CASE
                    WHEN EXISTS
                    (
                        SELECT 1
                        FROM dbo.CON_PlanCuenta AS h
                        WHERE h.IdEmpresa = pc.IdEmpresa
                          AND h.IdPlanCuentaPadre = pc.IdPlanCuenta
                          AND h.Estado = 1
                    )
                        THEN 0
                    ELSE 1
                END AS BIT) AS EsUltimoNivel,
                pc.RequiereCentroCosto,
                pc.Estado
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
              AND pc.Estado = 1
              AND (@SoloMovimiento = 0 )
              AND (@NivelCuenta IS NULL OR pc.NivelCuenta = @NivelCuenta)
              AND (
                    @SoloUltimoNivel = 0
                    OR NOT EXISTS
                    (
                        SELECT 1
                        FROM dbo.CON_PlanCuenta AS h
                        WHERE h.IdEmpresa = pc.IdEmpresa
                          AND h.IdPlanCuentaPadre = pc.IdPlanCuenta
                          AND h.Estado = 1
                    )
                  )
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR pc.CodigoCuenta LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR pc.NombreCuenta LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdPlanCuenta,
            b.IdPlanCuentaPadre,
            b.CodigoCuenta,
            b.NombreCuenta,
            b.NivelCuenta,
            b.ColBalance,
            b.IdMoneda,
            b.TipoCambio,
            b.AceptaMovimiento,
            b.EsUltimoNivel,
            b.RequiereCentroCosto,
            b.Estado,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.CodigoCuenta ASC
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
