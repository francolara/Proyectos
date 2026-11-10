-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_JornadaHYT_Listar para el registro de asistencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_JornadaHYT_Listar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Desde DATE = NULL,
    @Hasta DATE = NULL,
    @JornadaId BIGINT = NULL,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'ASISTENCIA.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar asistencia.', 1;
        IF @Pagina IS NULL OR @Pagina < 1 OR @Desde > @Hasta
            THROW 51000, N'Pagina o rango de fechas no valido.', 1;

        SELECT jornada.JornadaId, jornada.Entrada, jornada.Salida, jornada.MinutosExtrasAprobados,
               jornada.VersionFila, COUNT_BIG(1) OVER () AS Total,
               SUM(CONVERT(BIGINT, jornada.MinutosExtrasAprobados)) OVER () AS MinutosExtrasTotal
        FROM dbo.JornadaHYT AS jornada
        WHERE jornada.TrabajadorId = @TrabajadorId
          AND (@JornadaId IS NULL OR jornada.JornadaId = @JornadaId)
          AND (@Desde IS NULL OR jornada.Entrada >= @Desde)
          AND (@Hasta IS NULL OR CONVERT(DATE, jornada.Entrada) <= @Hasta)
        ORDER BY jornada.Entrada DESC, jornada.JornadaId DESC
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
