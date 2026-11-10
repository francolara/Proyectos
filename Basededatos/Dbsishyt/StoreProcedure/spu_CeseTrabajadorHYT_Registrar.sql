-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_CeseTrabajadorHYT_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CeseTrabajadorHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @TrabajadorId BIGINT,
    @FechaCese DATE,
    @UltimoDiaTrabajado DATE,
    @Motivo NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CESES') = 0
            THROW 51000, N'No tiene permiso para registrar ceses.', 1;

        SET @Motivo = NULLIF(TRIM(@Motivo), N'');
        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @OperacionId IS NULL OR @FechaCese IS NULL OR @UltimoDiaTrabajado IS NULL
           OR @UltimoDiaTrabajado > @FechaCese OR @FechaCese > @Hoy
            THROW 51000, N'Revise la fecha de cese y el ultimo dia trabajado.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        DECLARE @CeseId BIGINT;
        SELECT @CeseId = cese.CeseId FROM dbo.CeseTrabajadorHYT AS cese WHERE cese.OperacionId = @OperacionId;
        IF @CeseId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.CeseTrabajadorHYT AS cese
                WHERE cese.CeseId = @CeseId AND cese.TrabajadorId = @TrabajadorId
                  AND cese.FechaCese = @FechaCese AND cese.UltimoDiaTrabajado = @UltimoDiaTrabajado
                  AND cese.Motivo IS NOT DISTINCT FROM @Motivo)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            DECLARE @Mes DATE = DATEFROMPARTS(YEAR(@FechaCese), MONTH(@FechaCese), 1);
            IF EXISTS (SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes >= @Mes)
                THROW 51000, N'El cese corresponde a un mes cerrado o anterior al ultimo cierre.', 1;

            UPDATE dbo.TrabajadorHYT
            SET Activo = 0
            WHERE TrabajadorId = @TrabajadorId AND Activo = 1 AND FechaIngreso <= @UltimoDiaTrabajado;
            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador no esta activo o el ultimo dia es anterior al ingreso.', 1;

            INSERT INTO dbo.CeseTrabajadorHYT (TrabajadorId, FechaCese, UltimoDiaTrabajado, Motivo, OperacionId)
            VALUES (@TrabajadorId, @FechaCese, @UltimoDiaTrabajado, @Motivo, @OperacionId);

            SET @CeseId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CeseId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'CeseTrabajadorHYT',
                 @RegistroId, N'REGISTRAR', @EstadoNuevo = N'CESADO', @FechaEvento = @FechaCese, @Motivo = @Motivo;
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
