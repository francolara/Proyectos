-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_PeriodoVacacionesHYT_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PeriodoVacacionesHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @DiasReconocidos SMALLINT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar periodos de vacaciones.', 1;

        IF @FechaDesde IS NULL OR @FechaHasta IS NULL OR @FechaHasta < @FechaDesde
           OR @DiasReconocidos IS NULL OR @DiasReconocidos <= 0
            THROW 51000, N'Indique las fechas del periodo y los dias reconocidos.', 1;

        BEGIN TRANSACTION;
        DECLARE @FechaIngreso DATE;
        SELECT @FechaIngreso = trabajador.FechaIngreso
        FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
        WHERE trabajador.TrabajadorId = @TrabajadorId AND trabajador.Tipo = 'PLANILLA';

        IF @FechaIngreso IS NULL OR @FechaDesde < @FechaIngreso
            THROW 51000, N'El periodo debe corresponder a un trabajador en planilla y ser posterior a su ingreso o comenzar en esa fecha.', 1;

        IF EXISTS (
            SELECT 1 FROM dbo.PeriodoVacacionesHYT AS periodo
            WHERE periodo.TrabajadorId = @TrabajadorId
              AND periodo.FechaDesde <= @FechaHasta AND periodo.FechaHasta >= @FechaDesde)
            THROW 51000, N'El periodo se superpone con otro periodo registrado.', 1;

        INSERT INTO dbo.PeriodoVacacionesHYT (TrabajadorId, FechaDesde, FechaHasta, DiasReconocidos)
        VALUES (@TrabajadorId, @FechaDesde, @FechaHasta, @DiasReconocidos);

        DECLARE @PeriodoId BIGINT = SCOPE_IDENTITY();
        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PeriodoId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PeriodoVacacionesHYT', @RegistroId, N'REGISTRAR';

        COMMIT TRANSACTION;
        SELECT @PeriodoId AS PeriodoVacacionesId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
