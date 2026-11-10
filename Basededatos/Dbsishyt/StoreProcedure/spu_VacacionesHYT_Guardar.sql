-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_VacacionesHYT_Guardar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_VacacionesHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @PeriodoVacacionesId BIGINT,
    @FechaInicio DATE,
    @FechaFin DATE,
    @Dias SMALLINT,
    @Estado VARCHAR(10),
    @VacacionesId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar vacaciones.', 1;

        IF @FechaInicio IS NULL OR @FechaFin IS NULL OR @FechaFin < @FechaInicio
           OR @Dias IS NULL OR @Dias NOT BETWEEN 1 AND DATEDIFF(DAY, @FechaInicio, @FechaFin) + 1
           OR @Estado IS NULL OR @Estado NOT IN ('PENDIENTE', 'APROBADA', 'ANULADA')
           OR (@VacacionesId IS NULL AND @Estado = 'ANULADA')
            THROW 51000, N'Revise las fechas, los dias y el estado de las vacaciones.', 1;

        BEGIN TRANSACTION;
        DECLARE @FechaIngreso DATE;
        SELECT @FechaIngreso = trabajador.FechaIngreso
        FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
        WHERE trabajador.TrabajadorId = @TrabajadorId AND trabajador.Tipo = 'PLANILLA';

        IF @FechaIngreso IS NULL OR @FechaInicio < @FechaIngreso
            THROW 51000, N'Las vacaciones requieren un trabajador en planilla y fechas desde su ingreso.', 1;

        DECLARE @Reconocidos INT;
        SELECT @Reconocidos = periodo.DiasReconocidos
        FROM dbo.PeriodoVacacionesHYT AS periodo
        WHERE periodo.PeriodoVacacionesId = @PeriodoVacacionesId AND periodo.TrabajadorId = @TrabajadorId;

        IF @Reconocidos IS NULL
            THROW 51000, N'El periodo no corresponde al trabajador.', 1;

        DECLARE @EstadoAnterior VARCHAR(10);
        IF @VacacionesId IS NOT NULL
        BEGIN
            SELECT @EstadoAnterior = vacacion.Estado
            FROM dbo.VacacionesHYT AS vacacion
            WHERE vacacion.VacacionesId = @VacacionesId AND vacacion.PeriodoVacacionesId = @PeriodoVacacionesId
              AND vacacion.VersionFila = @VersionFila AND vacacion.Estado <> 'ANULADA';

            IF @EstadoAnterior IS NULL
                THROW 51000, N'El registro cambio o esta anulado. Vuelva a abrirlo.', 1;
        END;

        IF @Estado <> 'ANULADA' AND EXISTS (
            SELECT 1
            FROM dbo.VacacionesHYT AS vacacion
            INNER JOIN dbo.PeriodoVacacionesHYT AS periodo ON periodo.PeriodoVacacionesId = vacacion.PeriodoVacacionesId
            WHERE periodo.TrabajadorId = @TrabajadorId AND vacacion.Estado <> 'ANULADA'
              AND (@VacacionesId IS NULL OR vacacion.VacacionesId <> @VacacionesId)
              AND vacacion.FechaInicio <= @FechaFin AND vacacion.FechaFin >= @FechaInicio)
            THROW 51000, N'Las fechas se superponen con otras vacaciones registradas.', 1;

        IF @Estado = 'APROBADA'
        BEGIN
            DECLARE @Asignados INT;
            SELECT @Asignados = COALESCE(SUM(CONVERT(INT, vacacion.Dias)), 0)
            FROM dbo.VacacionesHYT AS vacacion
            WHERE vacacion.PeriodoVacacionesId = @PeriodoVacacionesId AND vacacion.Estado = 'APROBADA'
              AND (@VacacionesId IS NULL OR vacacion.VacacionesId <> @VacacionesId);

            IF @Dias > @Reconocidos - @Asignados
                THROW 51000, N'Los dias exceden el saldo disponible del periodo.', 1;
        END;

        IF @VacacionesId IS NULL
        BEGIN
            INSERT INTO dbo.VacacionesHYT (PeriodoVacacionesId, FechaInicio, FechaFin, Dias, Estado)
            VALUES (@PeriodoVacacionesId, @FechaInicio, @FechaFin, @Dias, @Estado);
            SET @VacacionesId = SCOPE_IDENTITY();
        END
        ELSE
            UPDATE dbo.VacacionesHYT
            SET FechaInicio = CASE WHEN @Estado = 'ANULADA' THEN FechaInicio ELSE @FechaInicio END,
                FechaFin = CASE WHEN @Estado = 'ANULADA' THEN FechaFin ELSE @FechaFin END,
                Dias = CASE WHEN @Estado = 'ANULADA' THEN Dias ELSE @Dias END,
                Estado = @Estado
            WHERE VacacionesId = @VacacionesId;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @VacacionesId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'VacacionesHYT', @RegistroId,
             N'GUARDAR', @EstadoAnterior, @Estado;

        COMMIT TRANSACTION;
        SELECT @VacacionesId AS VacacionesId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
