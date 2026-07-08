-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Elimina la generacion del asiento de apertura de un ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Elimina solo la generacion automatica APR del ejercicio, removiendo proceso, detalle, asiento asociado y correlativo del periodo 00 si queda vacio.

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarAperturaProceso
    @IdEmpresa INT,
    @AnioApertura SMALLINT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdOrigen INT
        DECLARE @PeriodoApertura CHAR(6) = CONCAT(@AnioApertura, '00')
        DECLARE @UltimoNumeroRestante INT = 0

        IF @AnioApertura < 2000 OR @AnioApertura > 9999
        BEGIN
            RAISERROR(N'El anio de apertura es invalido.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'APR'
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
            p.IdAsiento
        FROM dbo.CON_AperturaProceso AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.AnioApertura = @AnioApertura
          AND p.IdAsiento IS NOT NULL;

        DELETE d
        FROM dbo.CON_AperturaProcesoDetalle AS d
        INNER JOIN dbo.CON_AperturaProceso AS p
            ON p.IdAperturaProceso = d.IdAperturaProceso
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.AnioApertura = @AnioApertura;

        DELETE p
        FROM dbo.CON_AperturaProceso AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.AnioApertura = @AnioApertura;

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
            SELECT
                @UltimoNumeroRestante = ISNULL(MAX(a.NumeroAsiento), 0)
            FROM dbo.CON_Asiento AS a
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.IdOrigen = @IdOrigen
              AND a.Periodo = @PeriodoApertura;

            IF @UltimoNumeroRestante = 0
            BEGIN
                DELETE dbo.CON_CorrelativoAsiento
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @PeriodoApertura;
            END
            ELSE
            BEGIN
                UPDATE dbo.CON_CorrelativoAsiento
                SET UltimoNumero = @UltimoNumeroRestante,
                    FechaActualizacion = SYSDATETIME(),
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @PeriodoApertura;
            END;
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
