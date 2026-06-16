-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Genera el siguiente correlativo contable por empresa, origen y periodo mensual usando numerador seguro para concurrencia.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerSiguienteNumeroAsiento
    @IdEmpresa INT,
    @IdOrigen INT,
    @Periodo CHAR(6),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @SiguienteNumero INT
        DECLARE @Salida TABLE
        (
            SiguienteNumero INT NOT NULL
        );

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;

        BEGIN TRAN;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.IdOrigen = @IdOrigen
              AND c.Periodo = @Periodo
        )
        BEGIN
            UPDATE dbo.CON_CorrelativoAsiento
            SET UltimoNumero = UltimoNumero + 1,
                FechaActualizacion = SYSDATETIME(),
                UsuarioRegistro = @UsuarioRegistro
            OUTPUT inserted.UltimoNumero INTO @Salida (SiguienteNumero)
            WHERE IdEmpresa = @IdEmpresa
              AND IdOrigen = @IdOrigen
              AND Periodo = @Periodo;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.CON_CorrelativoAsiento
            (
                IdEmpresa,
                IdOrigen,
                Periodo,
                UltimoNumero,
                FechaActualizacion,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdOrigen,
                @Periodo,
                1,
                SYSDATETIME(),
                @UsuarioRegistro
            );

            INSERT INTO @Salida
            (
                SiguienteNumero
            )
            VALUES
            (
                1
            );
        END;

        SELECT TOP (1)
            @SiguienteNumero = s.SiguienteNumero
        FROM @Salida AS s;

        COMMIT;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        SELECT
            @IdEmpresa AS IdEmpresa,
            @IdOrigen AS IdOrigen,
            @Periodo AS Periodo,
            @SiguienteNumero AS NumeroAsiento;

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
