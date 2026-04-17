USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 16/04/2026 | Cancela automaticamente reservas no confirmadas segun check/minutos por negocio; apto para ejecucion directa desde SQL Job.
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_AutoCancelarNoConfirmadas
    @FechaHoraActual DATETIME2(7) = NULL,
    @Usuario NVARCHAR(120) = N'job_sql'
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Ahora DATETIME2(7);
        SET @Ahora = COALESCE(@FechaHoraActual, SYSDATETIME());

        DECLARE @Actualizadas TABLE (Id INT PRIMARY KEY);

        UPDATE r
            SET r.Estado = 5,
                r.FechaActualizacion = @Ahora,
                r.UsuarioActualizacion = @Usuario,
                r.Comentario = CASE
                    WHEN r.Comentario IS NULL OR LTRIM(RTRIM(r.Comentario)) = N''
                        THEN N'Reserva cancelada automaticamente por no confirmacion en el tiempo configurado.'
                    ELSE r.Comentario
                END
        OUTPUT INSERTED.Id INTO @Actualizadas(Id)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos ed ON ed.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = ed.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE r.Estado = 1
          AND n.CancelacionAutomaticaNoConfirmada = 1
          AND COALESCE(n.MinutosCancelacionNoConfirmada, 30) > 0
          AND DATEDIFF(MINUTE, r.FechaRegistro, @Ahora) >= COALESCE(n.MinutosCancelacionNoConfirmada, 30)
          AND NOT EXISTS (
              SELECT 1
              FROM dbo.Pagos p
              WHERE p.ReservaId = r.Id
          );

        SELECT COUNT(1)
        FROM @Actualizadas;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
