-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Calcula saldo inicial, ingresos, egresos y saldo final del periodo por cuenta corriente corrigiendo el agregado sobre la CTE Base.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_BAN_ObtenerResumenMovimientoBanco
    @IdEmpresa INT,
    @IdBancoConfiguracionEmpresa INT = NULL,
    @Anio SMALLINT,
    @Mes TINYINT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @FechaInicio DATE = DATEFROMPARTS(@Anio, @Mes, 1);
        DECLARE @FechaFin DATE = DATEADD(MONTH, 1, @FechaInicio);

        ;WITH Base AS
        (
            SELECT
                m.TipoMovimiento,
                m.ImporteTotal,
                m.FechaEmision
            FROM dbo.BAN_MovimientoBanco AS m
            WHERE m.IdEmpresa = @IdEmpresa
              AND m.Activo = 1
              AND (@IdBancoConfiguracionEmpresa IS NULL OR m.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa)
        )
        SELECT
            CAST(ISNULL(SUM(CASE WHEN b.FechaEmision < @FechaInicio AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0)
               - ISNULL(SUM(CASE WHEN b.FechaEmision < @FechaInicio AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0) AS DECIMAL(18, 2)) AS SaldoInicial,
            CAST(ISNULL(SUM(CASE WHEN b.FechaEmision >= @FechaInicio AND b.FechaEmision < @FechaFin AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0) AS DECIMAL(18, 2)) AS IngresosMes,
            CAST(ISNULL(SUM(CASE WHEN b.FechaEmision >= @FechaInicio AND b.FechaEmision < @FechaFin AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0) AS DECIMAL(18, 2)) AS EgresosMes,
            CAST(
                (
                    ISNULL(SUM(CASE WHEN b.FechaEmision < @FechaInicio AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0)
                  - ISNULL(SUM(CASE WHEN b.FechaEmision < @FechaInicio AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0)
                  + ISNULL(SUM(CASE WHEN b.FechaEmision >= @FechaInicio AND b.FechaEmision < @FechaFin AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0)
                  - ISNULL(SUM(CASE WHEN b.FechaEmision >= @FechaInicio AND b.FechaEmision < @FechaFin AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0)
                ) AS DECIMAL(18, 2)
            ) AS SaldoFinal
        FROM Base AS b;

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
