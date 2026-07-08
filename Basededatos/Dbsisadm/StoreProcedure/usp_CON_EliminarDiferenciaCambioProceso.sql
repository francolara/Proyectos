-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Elimina la generacion de diferencia en cambio de un periodo.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Elimina solo la generacion automatica DIF del periodo, removiendo proceso, detalle, asientos asociados y correlativo sobrante.

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarDiferenciaCambioProceso
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdOrigen INT
        DECLARE @UltimoNumeroRestante INT = 0

        IF @Periodo IS NULL
           OR @Periodo NOT LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
           OR RIGHT(@Periodo, 2) NOT BETWEEN '01' AND '12'
        BEGIN
            RAISERROR(N'El periodo debe estar en formato yyyyMM.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'DIF'
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
        FROM dbo.CON_DiferenciaCambioProcesoDetalle AS d
        INNER JOIN dbo.CON_DiferenciaCambioProceso AS p
            ON p.IdDiferenciaCambioProceso = d.IdDiferenciaCambioProceso
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Periodo = @Periodo
          AND d.IdAsiento IS NOT NULL;

        DELETE d
        FROM dbo.CON_DiferenciaCambioProcesoDetalle AS d
        INNER JOIN dbo.CON_DiferenciaCambioProceso AS p
            ON p.IdDiferenciaCambioProceso = d.IdDiferenciaCambioProceso
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Periodo = @Periodo;

        DELETE p
        FROM dbo.CON_DiferenciaCambioProceso AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Periodo = @Periodo;

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
              AND a.Periodo = @Periodo;

            IF @UltimoNumeroRestante = 0
            BEGIN
                DELETE dbo.CON_CorrelativoAsiento
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @Periodo;
            END
            ELSE
            BEGIN
                UPDATE dbo.CON_CorrelativoAsiento
                SET UltimoNumero = @UltimoNumeroRestante,
                    FechaActualizacion = SYSDATETIME(),
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @Periodo;
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
