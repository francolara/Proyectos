USE [DbSportCenter]
GO
/****** Object:  StoredProcedure [dbo].[Sp_Sedes_ObtenerPorId]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 24_Sedes_Horario_NoLaborable.sql (linea 104)
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Devuelve ConsideracionesReserva para mantenimiento de sedes.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Devuelve URLs sociales (Facebook/Instagram/Twitter) para mantenimiento de sedes.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Sedes_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id, s.NegocioId, s.Nombre, s.Direccion, s.ConsideracionesReserva, s.Telefono,
            s.FacebookUrl, s.InstagramUrl, s.TwitterUrl,
            s.Activo,
            s.Latitud,
            s.Longitud,
            s.GooglePlaceId,
            s.GoogleMapsUrl,
            s.FotoPrincipalUrl,
            s.FotosUrlsCsv,
            STUFF((SELECT N',' + CONVERT(NVARCHAR(20), ss.ServicioId) FROM dbo.SedeServicios ss WHERE ss.SedeId = s.Id ORDER BY ss.ServicioId FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS ServiciosIdsCsv,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivas,
            COALESCE(scn.MinutosAnticipacionRecordatorio, 90) AS MinutosAnticipacionRecordatorio,
            COALESCE(scn.MinutosToleranciaNoShow, 30) AS MinutosToleranciaNoShow,
            scn.CorreoNotificacion,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            COALESCE(sha.AtiendeLunes, 1) AS AtiendeLunes,
            COALESCE(sha.AtiendeMartes, 1) AS AtiendeMartes,
            COALESCE(sha.AtiendeMiercoles, 1) AS AtiendeMiercoles,
            COALESCE(sha.AtiendeJueves, 1) AS AtiendeJueves,
            COALESCE(sha.AtiendeViernes, 1) AS AtiendeViernes,
            COALESCE(sha.AtiendeSabado, 1) AS AtiendeSabado,
            COALESCE(sha.AtiendeDomingo, 1) AS AtiendeDomingo,
            COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) AS HoraApertura,
            COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) AS HoraCierre,
            STUFF((SELECT N',' + CONVERT(NVARCHAR(10), sfi.Fecha, 23) FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = s.Id AND sfi.Activo = 1 ORDER BY sfi.Fecha FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS FechasInhabilitadasCsv
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND s.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
