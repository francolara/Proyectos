USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/04/2026
-- Description:   Obtiene contexto de reserva para correos (cliente/sede/notificaciones/negocio).
-- Firma:         Codex - 26/04/2026 | Nuevo SP para notificacion de reserva publica y confirmacion. Incluye nombre del club/negocio para cuerpo de correo.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_ObtenerContextoEmail]
    @NegocioId INT = NULL,
    @ReservaId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT TOP (1)
            r.Id AS ReservaId,
            s.NegocioId,
            n.NombreComercial AS Negocio,
            r.Estado,
            c.NombresORazonSocial AS Cliente,
            c.Correo AS ClienteCorreo,
            c.Telefono AS ClienteTelefono,
            c.NombreEquipo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivasSede,
            scn.CorreoNotificacion AS CorreoNotificacionSede
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE r.Id = @ReservaId
          AND (@NegocioId IS NULL OR s.NegocioId = @NegocioId);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
