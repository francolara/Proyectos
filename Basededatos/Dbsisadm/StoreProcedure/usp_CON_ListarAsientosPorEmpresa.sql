-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista los asientos contables por empresa con filtro por periodo, busqueda y paginacion server-side.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarAsientosPorEmpresa
    @IdEmpresa INT,
    @Periodo CHAR(6) = NULL,
    @Ejercicio SMALLINT = NULL,
    @Mes TINYINT = NULL,
    @SoloManual BIT = 0,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @PeriodoTrabajo CHAR(6) =
            CASE
                WHEN @Periodo IS NOT NULL THEN @Periodo
                WHEN @Ejercicio IS NOT NULL AND @Mes IS NOT NULL THEN CONVERT(CHAR(4), @Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), @Mes), 2)
                ELSE NULL
            END
        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                a.IdAsiento,
                a.IdEmpresa,
                a.IdOrigen,
                o.CodigoOrigen,
                o.NombreOrigen,
                o.PermiteRegistroManual,
                a.Ejercicio,
                a.Mes,
                a.Periodo,
                a.NumeroAsiento,
                a.FechaAsiento,
                a.Glosa,
                a.IdMoneda,
                m.CodigoMoneda,
                m.NombreMoneda,
                m.SimboloMoneda,
                a.TipoCambio,
                a.TotalDebe,
                a.TotalHaber,
                a.Estado,
                a.ReferenciaExterna,
                a.Observacion
            FROM dbo.CON_Asiento AS a
            INNER JOIN dbo.CON_Origen AS o
                ON o.IdOrigen = a.IdOrigen
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = a.IdMoneda
            WHERE a.IdEmpresa = @IdEmpresa
              AND (@PeriodoTrabajo IS NULL OR a.Periodo = @PeriodoTrabajo)
              AND (@SoloManual = 0 OR o.PermiteRegistroManual = 1)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR o.CodigoOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR o.NombreOrigen LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR a.Glosa LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(a.ReferenciaExterna, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR CONVERT(VARCHAR(20), a.NumeroAsiento) LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdAsiento,
            b.IdEmpresa,
            b.IdOrigen,
            b.CodigoOrigen,
            b.NombreOrigen,
            b.PermiteRegistroManual,
            b.Ejercicio,
            b.Mes,
            b.Periodo,
            b.NumeroAsiento,
            b.FechaAsiento,
            b.Glosa,
            b.IdMoneda,
            b.CodigoMoneda,
            b.NombreMoneda,
            b.SimboloMoneda,
            b.TipoCambio,
            b.TotalDebe,
            b.TotalHaber,
            b.Estado,
            b.ReferenciaExterna,
            b.Observacion,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY
            b.FechaAsiento DESC,
            b.NumeroAsiento DESC,
            b.IdAsiento DESC
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
