USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Lista historial de reservas realizadas por usuario publico autenticado.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/04/2026
-- Description:   Incluye URLs de redes y mapa de la sede para vista en tarjetas del perfil publico.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_UsuariosPublicos_ReservasListar]
    @UsuarioId NVARCHAR(450),
    @Top INT = 200
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @TopFinal INT = CASE WHEN @Top IS NULL OR @Top <= 0 THEN 200 ELSE @Top END;

        SELECT TOP (@TopFinal)
            r.Id AS ReservaId,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            CASE r.Estado
                WHEN 1 THEN N'Reservada'
                WHEN 2 THEN N'Confirmada'
                WHEN 3 THEN N'Pagada'
                WHEN 4 THEN N'Completada'
                WHEN 5 THEN N'Cancelada'
                WHEN 6 THEN N'No Show'
                ELSE N'Pendiente'
            END AS EstadoTexto,
            CAST(r.Total AS DECIMAL(10,2)) AS Total,
            CAST(r.Adelanto AS DECIMAL(10,2)) AS Adelanto,
            CAST((r.Total - r.Adelanto) AS DECIMAL(10,2)) AS SaldoPendiente,
            n.NombreComercial AS NegocioNombre,
            s.Nombre AS SedeNombre,
            e.Nombre AS EspacioNombre,
            s.Direccion AS SedeDireccion,
            s.Telefono AS SedeTelefono,
            scn.WhatsappContacto AS SedeWhatsapp,
            s.FacebookUrl AS SedeFacebookUrl,
            s.InstagramUrl AS SedeInstagramUrl,
            s.TwitterUrl AS SedeTwitterUrl,
            COALESCE(
                NULLIF(LTRIM(RTRIM(s.GoogleMapsUrl)), N''),
                CASE
                    WHEN s.Latitud IS NOT NULL AND s.Longitud IS NOT NULL
                        THEN N'https://www.google.com/maps?q='
                            + CONVERT(NVARCHAR(40), s.Latitud)
                            + N','
                            + CONVERT(NVARCHAR(40), s.Longitud)
                    ELSE NULL
                END
            ) AS SedeMapaUrl
        FROM dbo.ReservasUsuariosPublicos rup
        INNER JOIN dbo.Reservas r ON r.Id = rup.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE rup.UsuarioId = @UsuarioId
        ORDER BY r.Fecha DESC, r.HoraInicio DESC, r.Id DESC;
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
