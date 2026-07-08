-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Calcula saldo inicial, ingresos, egresos y saldo final del periodo por cuenta corriente corrigiendo el agregado sobre la CTE Base.
-- =============================================
-- Firma: FRANCO LARA - 03/07/2026 | Ajusta el resumen de Caja y Bancos para cortar saldos e importes por el Periodo persistido de cada movimiento.

CREATE OR ALTER PROCEDURE dbo.usp_BAN_ObtenerResumenMovimientoBanco
    @IdEmpresa INT,
    @IdBancoConfiguracionEmpresa INT = NULL,
    @Anio SMALLINT,
    @Mes TINYINT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), @Anio) + RIGHT('0' + CONVERT(VARCHAR(2), @Mes), 2);

        ;WITH Base AS
        (
            SELECT
                m.TipoMovimiento,
                m.ImporteTotal,
                m.Periodo
            FROM dbo.BAN_MovimientoBanco AS m
            WHERE m.IdEmpresa = @IdEmpresa
              AND m.Activo = 1
              AND (@IdBancoConfiguracionEmpresa IS NULL OR m.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa)
        )
        SELECT
            CAST(ISNULL(SUM(CASE WHEN b.Periodo < @Periodo AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0)
               - ISNULL(SUM(CASE WHEN b.Periodo < @Periodo AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0) AS DECIMAL(18, 2)) AS SaldoInicial,
            CAST(ISNULL(SUM(CASE WHEN b.Periodo = @Periodo AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0) AS DECIMAL(18, 2)) AS IngresosMes,
            CAST(ISNULL(SUM(CASE WHEN b.Periodo = @Periodo AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0) AS DECIMAL(18, 2)) AS EgresosMes,
            CAST(
                (
                    ISNULL(SUM(CASE WHEN b.Periodo < @Periodo AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0)
                  - ISNULL(SUM(CASE WHEN b.Periodo < @Periodo AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0)
                  + ISNULL(SUM(CASE WHEN b.Periodo = @Periodo AND b.TipoMovimiento = 'I' THEN b.ImporteTotal END), 0)
                  - ISNULL(SUM(CASE WHEN b.Periodo = @Periodo AND b.TipoMovimiento = 'E' THEN b.ImporteTotal END), 0)
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
