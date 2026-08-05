-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Obtiene la ultima generacion, continuidad mensual y snapshot presentado del ejercicio.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_Presentacion_ObtenerContexto
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @PeriodoAnterior CHAR(6),
    @CodigoLibro VARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @IdActual INT;
        DECLARE @Presentado BIT = 0;
        DECLARE @FechaPresentacion DATETIME2(0);
        DECLARE @UsuarioPresentacion NVARCHAR(450);

        SELECT TOP (1)
            @IdActual = h.IdLibroElectronicoGeneracion,
            @Presentado = h.PlanPresentado,
            @FechaPresentacion = h.FechaPresentacion,
            @UsuarioPresentacion = h.UsuarioPresentacion
        FROM dbo.CON_LibroElectronicoGeneracion AS h
        WHERE h.IdEmpresa = @IdEmpresa
          AND h.Periodo = @Periodo
          AND h.CodigoLibro = @CodigoLibro
        ORDER BY h.IdLibroElectronicoGeneracion DESC;

        SELECT
            @IdActual AS IdGeneracionPeriodo,
            ISNULL(@Presentado, 0) AS Presentado,
            @FechaPresentacion AS FechaPresentacion,
            @UsuarioPresentacion AS UsuarioPresentacion,
            CONVERT(BIT, CASE WHEN EXISTS
            (
                SELECT 1
                FROM dbo.CON_LibroElectronicoGeneracion AS h
                WHERE h.IdEmpresa = @IdEmpresa
                  AND h.CodigoLibro = @CodigoLibro
                  AND h.Periodo < @Periodo
                  AND h.PlanPresentado = 1
            ) THEN 1 ELSE 0 END) AS ExistePresentacionAnterior,
            CONVERT(BIT, CASE WHEN EXISTS
            (
                SELECT 1
                FROM dbo.CON_LibroElectronicoGeneracion AS h
                WHERE h.IdEmpresa = @IdEmpresa
                  AND h.CodigoLibro = @CodigoLibro
                  AND h.Periodo = @PeriodoAnterior
                  AND h.PlanPresentado = 1
            ) THEN 1 ELSE 0 END) AS MesAnteriorPresentado,
            CONVERT(BIT, CASE WHEN EXISTS
            (
                SELECT 1
                FROM dbo.CON_LibroElectronicoGeneracion AS h
                WHERE h.IdEmpresa = @IdEmpresa
                  AND h.CodigoLibro = @CodigoLibro
                  AND h.Periodo > @Periodo
                  AND h.PlanPresentado = 1
            ) THEN 1 ELSE 0 END) AS ExistePresentacionPosterior,
            (
                SELECT TOP (1) h.PlanContableSnapshot
                FROM dbo.CON_LibroElectronicoGeneracion AS h
                WHERE h.IdEmpresa = @IdEmpresa
                  AND h.CodigoLibro = @CodigoLibro
                  AND h.Periodo < @Periodo
                  AND LEFT(h.Periodo, 4) = LEFT(@Periodo, 4)
                  AND h.PlanPresentado = 1
                  AND h.PlanContableSnapshot IS NOT NULL
                ORDER BY h.Periodo DESC, h.IdLibroElectronicoGeneracion DESC
            ) AS SnapshotUltimaPresentacion;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
