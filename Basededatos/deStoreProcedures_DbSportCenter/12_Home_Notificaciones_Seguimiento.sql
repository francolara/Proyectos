-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Soporte de seguimiento publico por codigo y notificacion de solicitudes.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Home_ConsultarSolicitudPublica
    @CodigoSolicitud NVARCHAR(20),
    @Telefono NVARCHAR(30)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (1)
            s.CodigoSolicitud,
            se.Nombre AS Sede,
            e.Nombre AS Espacio,
            s.Fecha,
            s.HoraInicio,
            s.HoraFin,
            s.NombreSolicitante,
            s.Telefono,
            s.Correo,
            s.Estado,
            CASE s.Estado
                WHEN 1 THEN N'Pendiente'
                WHEN 2 THEN N'Aprobada'
                WHEN 3 THEN N'Rechazada'
                WHEN 4 THEN N'Convertida a reserva'
                ELSE N'Desconocido'
            END AS EstadoTexto,
            s.ReservaId,
            s.FechaRegistro
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.CodigoSolicitud = @CodigoSolicitud
          AND s.Telefono = @Telefono
        ORDER BY s.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_ObtenerSolicitudParaNotificacion
    @CodigoSolicitud NVARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (1)
            s.CodigoSolicitud,
            s.NombreSolicitante,
            s.Correo,
            se.Nombre AS Sede,
            e.Nombre AS Espacio,
            s.Fecha,
            s.HoraInicio,
            s.HoraFin,
            s.NotificadoCliente
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.CodigoSolicitud = @CodigoSolicitud;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_MarcarSolicitudNotificada
    @CodigoSolicitud NVARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.SolicitudesReservaPublica
        SET NotificadoCliente = 1
        WHERE CodigoSolicitud = @CodigoSolicitud;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
