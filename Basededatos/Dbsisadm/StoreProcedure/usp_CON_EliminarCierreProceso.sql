-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Elimina la generacion del asiento de cierre de un ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Elimina solo la generacion automatica CIE del ejercicio, removiendo proceso, detalle, asientos asociados y correlativos vacios de los periodos 14 y 15.

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarCierreProceso
    @IdEmpresa INT,
    @Anio SMALLINT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdOrigen INT
        DECLARE @PeriodoGanancias CHAR(6) = CONCAT(@Anio, '14')
        DECLARE @PeriodoInventarios CHAR(6) = CONCAT(@Anio, '15')

        IF @Anio < 2000 OR @Anio > 9999
        BEGIN
            RAISERROR(N'El ejercicio indicado es invalido.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'CIE'
          AND c.EscenarioOperacion = 'PROVISION'
          AND c.Activo = 1;

        DECLARE @AsientosEliminar TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        INSERT INTO @AsientosEliminar (IdAsiento)
        SELECT DISTINCT
            d.IdAsiento
        FROM dbo.CON_CierreProcesoDetalle AS d
        INNER JOIN dbo.CON_CierreProceso AS p
            ON p.IdCierreProceso = d.IdCierreProceso
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Anio = @Anio
          AND d.IdAsiento IS NOT NULL;

        DELETE d
        FROM dbo.CON_CierreProcesoDetalle AS d
        INNER JOIN dbo.CON_CierreProceso AS p
            ON p.IdCierreProceso = d.IdCierreProceso
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Anio = @Anio;

        DELETE p
        FROM dbo.CON_CierreProceso AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Anio = @Anio;

        DELETE d
        FROM dbo.CON_AsientoDetalle AS d
        INNER JOIN @AsientosEliminar AS e
            ON e.IdAsiento = d.IdAsiento;

        DELETE a
        FROM dbo.CON_Asiento AS a
        INNER JOIN @AsientosEliminar AS e
            ON e.IdAsiento = a.IdAsiento;

        IF @IdOrigen IS NOT NULL
        BEGIN
            DELETE correlativo
            FROM dbo.CON_CorrelativoAsiento AS correlativo
            WHERE correlativo.IdEmpresa = @IdEmpresa
              AND correlativo.IdOrigen = @IdOrigen
              AND correlativo.Periodo IN (@PeriodoGanancias, @PeriodoInventarios)
              AND NOT EXISTS
              (
                  SELECT 1
                  FROM dbo.CON_Asiento AS a
                  WHERE a.IdEmpresa = @IdEmpresa
                    AND a.IdOrigen = @IdOrigen
                    AND a.Periodo = correlativo.Periodo
              );

            UPDATE correlativo
            SET UltimoNumero = base.UltimoNumero,
                FechaActualizacion = SYSDATETIME(),
                UsuarioRegistro = @UsuarioRegistro
            FROM dbo.CON_CorrelativoAsiento AS correlativo
            INNER JOIN
            (
                SELECT
                    a.Periodo,
                    MAX(a.NumeroAsiento) AS UltimoNumero
                FROM dbo.CON_Asiento AS a
                WHERE a.IdEmpresa = @IdEmpresa
                  AND a.IdOrigen = @IdOrigen
                  AND a.Periodo IN (@PeriodoGanancias, @PeriodoInventarios)
                GROUP BY
                    a.Periodo
            ) AS base
                ON base.Periodo = correlativo.Periodo
            WHERE correlativo.IdEmpresa = @IdEmpresa
              AND correlativo.IdOrigen = @IdOrigen;
        END;

        COMMIT;
        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

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
