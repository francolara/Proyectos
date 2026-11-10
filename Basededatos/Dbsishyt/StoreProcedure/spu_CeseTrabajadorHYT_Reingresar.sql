-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Registra el reingreso, conserva el cese y habilita al trabajador.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CeseTrabajadorHYT_Reingresar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @TrabajadorId BIGINT,
    @CeseId BIGINT,
    @FechaReingreso DATE,
    @Motivo NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CESES') = 0
            THROW 51000, N'No tiene permiso para registrar reingresos.', 1;

        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @OperacionId IS NULL OR @FechaReingreso IS NULL OR @FechaReingreso > @Hoy
            THROW 51000, N'Revise la fecha de reingreso.', 1;
        SET @Motivo = NULLIF(TRIM(@Motivo), N'');

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        IF NOT EXISTS (
            SELECT 1 FROM dbo.CeseTrabajadorHYT AS cese
            WHERE cese.CeseId = @CeseId AND cese.TrabajadorId = @TrabajadorId
              AND cese.OperacionReingreso = @OperacionId AND cese.FechaReingreso = @FechaReingreso
              AND cese.MotivoReingreso IS NOT DISTINCT FROM @Motivo)
        BEGIN
            DECLARE @UltimoCeseId BIGINT;
            SELECT TOP (1) @UltimoCeseId = cese.CeseId
            FROM dbo.CeseTrabajadorHYT AS cese
            WHERE cese.TrabajadorId = @TrabajadorId
            ORDER BY cese.FechaCese DESC, cese.CeseId DESC;

            UPDATE dbo.TrabajadorHYT
            SET Activo = 1, FechaIngreso = @FechaReingreso
            WHERE TrabajadorId = @TrabajadorId AND Activo = 0;
            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador no esta inactivo.', 1;

            UPDATE dbo.CeseTrabajadorHYT
            SET FechaReingreso = @FechaReingreso, MotivoReingreso = @Motivo, OperacionReingreso = @OperacionId
            WHERE CeseId = @CeseId AND CeseId = @UltimoCeseId AND TrabajadorId = @TrabajadorId
              AND FechaReingreso IS NULL AND FechaCese < @FechaReingreso;
            IF @@ROWCOUNT = 0
                THROW 51000, N'El cese no esta vigente o la fecha de reingreso no es posterior al cese.', 1;

            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CeseId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'CeseTrabajadorHYT',
                 @RegistroId, N'REINGRESAR', @EstadoAnterior = N'CESADO', @EstadoNuevo = N'ACTIVO',
                 @FechaEvento = @FechaReingreso, @Motivo = @Motivo;
        END;

        COMMIT TRANSACTION;
        SELECT @CeseId AS CeseId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
