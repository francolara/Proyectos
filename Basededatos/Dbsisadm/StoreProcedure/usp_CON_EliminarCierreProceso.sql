-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Elimina la generacion del asiento de cierre de un ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Elimina solo la generacion automatica CIE del ejercicio, removiendo proceso, detalle, asientos asociados y correlativos vacios de los periodos 14 y 15.
-- Firma: FRANCO LARA - 13/08/2026 | Elimina exclusivamente los asientos vinculados al cierre y recompone los correlativos de todos los periodos afectados, incluyendo generaciones heredadas.

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarCierreProceso
    @IdEmpresa INT,
    @Anio SMALLINT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF @Anio < 2000 OR @Anio > 9999
        BEGIN
            RAISERROR(N'El ejercicio indicado es invalido.', 16, 1);
        END;

        DECLARE @AsientosEliminar TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        DECLARE @CorrelativosRecalcular TABLE
        (
            IdOrigen INT NOT NULL,
            Periodo CHAR(6) NOT NULL,
            PRIMARY KEY (IdOrigen, Periodo)
        );

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        INSERT INTO @AsientosEliminar (IdAsiento)
        SELECT DISTINCT
            x.IdAsiento
        FROM
        (
            SELECT p.IdAsiento
            FROM dbo.CON_CierreProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Anio = @Anio
              AND p.IdAsiento IS NOT NULL

            UNION

            SELECT d.IdAsiento
            FROM dbo.CON_CierreProcesoDetalle AS d
            INNER JOIN dbo.CON_CierreProceso AS p
                ON p.IdCierreProceso = d.IdCierreProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Anio = @Anio
              AND d.IdAsiento IS NOT NULL
        ) AS x;

        INSERT INTO @CorrelativosRecalcular (IdOrigen, Periodo)
        SELECT DISTINCT
            a.IdOrigen,
            a.Periodo
        FROM dbo.CON_Asiento AS a
        INNER JOIN @AsientosEliminar AS e
            ON e.IdAsiento = a.IdAsiento;

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

        UPDATE correlativo
        SET UltimoNumero = base.UltimoNumero,
            FechaActualizacion = SYSDATETIME(),
            UsuarioRegistro = @UsuarioRegistro
        FROM dbo.CON_CorrelativoAsiento AS correlativo
        INNER JOIN @CorrelativosRecalcular AS r
            ON r.IdOrigen = correlativo.IdOrigen
           AND r.Periodo = correlativo.Periodo
        INNER JOIN
        (
            SELECT
                a.IdOrigen,
                a.Periodo,
                MAX(a.NumeroAsiento) AS UltimoNumero
            FROM dbo.CON_Asiento AS a
            INNER JOIN @CorrelativosRecalcular AS r
                ON r.IdOrigen = a.IdOrigen
               AND r.Periodo = a.Periodo
            WHERE a.IdEmpresa = @IdEmpresa
            GROUP BY
                a.IdOrigen,
                a.Periodo
        ) AS base
            ON base.IdOrigen = correlativo.IdOrigen
           AND base.Periodo = correlativo.Periodo
        WHERE correlativo.IdEmpresa = @IdEmpresa;

        DELETE correlativo
        FROM dbo.CON_CorrelativoAsiento AS correlativo
        INNER JOIN @CorrelativosRecalcular AS r
            ON r.IdOrigen = correlativo.IdOrigen
           AND r.Periodo = correlativo.Periodo
        WHERE correlativo.IdEmpresa = @IdEmpresa
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_Asiento AS a
              WHERE a.IdEmpresa = @IdEmpresa
                AND a.IdOrigen = correlativo.IdOrigen
                AND a.Periodo = correlativo.Periodo
          );

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
