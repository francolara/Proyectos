USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 16/04/2026 | Marca una notificacion como leida para el negocio.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Notificaciones_MarcarLeida]
    @NegocioId INT,
    @NotificacionId INT,
    @UserId NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        UPDATE n
           SET n.Leida = 1,
               n.FechaLeidaUtc = SYSUTCDATETIME(),
               n.LeidaPorUserId = NULLIF(LTRIM(RTRIM(@UserId)), N'')
        FROM dbo.NegocioNotificaciones n
        WHERE n.NegocioId = @NegocioId
          AND n.Id = @NotificacionId
          AND n.Leida = 0;

        SELECT @@ROWCOUNT AS FilasAfectadas;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
