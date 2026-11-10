-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_CierrePersonalHYT_Registrar para pagos del personal.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CierrePersonalHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CERRAR') = 0
            THROW 51000, N'No tiene permiso para cerrar el mes.', 1;

        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @Mes IS NULL OR DAY(@Mes) <> 1 OR EOMONTH(@Mes) >= @Hoy
            THROW 51000, N'Seleccione un mes que ya haya terminado.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        IF NOT EXISTS (SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes = @Mes)
        BEGIN
            IF EXISTS (
                SELECT 1 FROM dbo.MovimientoPersonalHYT AS movimiento
                LEFT JOIN dbo.CierrePersonalHYT AS cierre
                  ON cierre.Mes = DATEFROMPARTS(YEAR(movimiento.Fecha), MONTH(movimiento.Fecha), 1)
                WHERE movimiento.Fecha < @Mes AND cierre.Mes IS NULL)
                THROW 51000, N'Cierre primero los meses anteriores que tienen movimientos.', 1;

            INSERT INTO dbo.CierrePersonalHYT (Mes, UsuarioId) VALUES (@Mes, @UsuarioId);

            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(10), @Mes, 23);
            DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'CierrePersonalHYT',
                 @RegistroId, N'CERRAR', @EstadoNuevo = N'CERRADO';
        END;

        COMMIT TRANSACTION;
        SELECT @Mes AS Mes;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
