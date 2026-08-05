-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Registra o actualiza la huella del plan PLE generado por empresa, ejercicio y formato.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_PlanContable_ControlGuardar
    @IdEmpresa INT,
    @Anio SMALLINT,
    @CodigoFormato VARCHAR(10),
    @HuellaPlanContable CHAR(64),
    @UsuarioGeneracion NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        BEGIN TRANSACTION;

        UPDATE pc WITH (UPDLOCK, SERIALIZABLE)
        SET
            pc.HuellaPlanContable = @HuellaPlanContable,
            pc.FechaUltimaGeneracion = SYSDATETIME(),
            pc.UsuarioGeneracion = @UsuarioGeneracion
        FROM dbo.CON_PLE_PlanContableControl AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.Anio = @Anio
          AND pc.CodigoFormato = @CodigoFormato;

        IF @@ROWCOUNT = 0
        BEGIN
            INSERT INTO dbo.CON_PLE_PlanContableControl
            (
                IdEmpresa,
                Anio,
                CodigoFormato,
                HuellaPlanContable,
                FechaUltimaGeneracion,
                UsuarioGeneracion
            )
            VALUES
            (
                @IdEmpresa,
                @Anio,
                @CodigoFormato,
                @HuellaPlanContable,
                SYSDATETIME(),
                @UsuarioGeneracion
            );
        END;

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF XACT_STATE() <> 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

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
