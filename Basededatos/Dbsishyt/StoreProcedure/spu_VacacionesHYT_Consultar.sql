-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_VacacionesHYT_Consultar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_VacacionesHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Hoy DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar vacaciones.', 1;

        SELECT periodo.PeriodoVacacionesId, periodo.FechaDesde, periodo.FechaHasta,
               periodo.DiasReconocidos, periodo.VersionFila,
               COALESCE(saldo.Gozados, 0) AS DiasGozados,
               COALESCE(saldo.Programados, 0) AS DiasProgramados,
               periodo.DiasReconocidos - COALESCE(saldo.Gozados, 0) - COALESCE(saldo.Programados, 0) AS DiasDisponibles
        FROM dbo.PeriodoVacacionesHYT AS periodo
        OUTER APPLY (
            SELECT SUM(CASE WHEN vacacion.FechaFin < @Hoy THEN CONVERT(INT, vacacion.Dias) ELSE 0 END) AS Gozados,
                   SUM(CASE WHEN vacacion.FechaFin >= @Hoy THEN CONVERT(INT, vacacion.Dias) ELSE 0 END) AS Programados
            FROM dbo.VacacionesHYT AS vacacion
            WHERE vacacion.PeriodoVacacionesId = periodo.PeriodoVacacionesId AND vacacion.Estado = 'APROBADA'
        ) AS saldo
        WHERE periodo.TrabajadorId = @TrabajadorId
        ORDER BY periodo.FechaDesde DESC;

        SELECT vacacion.VacacionesId, vacacion.PeriodoVacacionesId, vacacion.FechaInicio,
               vacacion.FechaFin, vacacion.Dias, vacacion.Estado, vacacion.VersionFila
        FROM dbo.VacacionesHYT AS vacacion
        INNER JOIN dbo.PeriodoVacacionesHYT AS periodo ON periodo.PeriodoVacacionesId = vacacion.PeriodoVacacionesId
        WHERE periodo.TrabajadorId = @TrabajadorId
        ORDER BY vacacion.FechaInicio DESC, vacacion.VacacionesId DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
