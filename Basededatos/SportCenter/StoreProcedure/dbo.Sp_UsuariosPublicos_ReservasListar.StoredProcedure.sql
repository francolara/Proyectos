
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
-- Firma:         Codex - 26/04/2026 | Se implementa paginacion real por pagina/tamano (6 por defecto) desde SQL para Mis Reservas.
-- Firma:         FRANCO LARA - 08/06/2026 | Se incorpora estado de reseña por reserva para permitir un solo registro en estados Confirmada, Pagada o Completada y mostrar la reseña existente en el perfil publico.
-- Firma:         FRANCO LARA - 11/06/2026 | Devuelve estado visible y respuesta administrativa de cada reseña para reflejar la gestion del negocio en Mis reservas.
CREATE OR ALTER PROCEDURE [dbo].[Sp_UsuariosPublicos_ReservasListar]
    @UsuarioId NVARCHAR(450),
    @Pagina INT = 1,
    @TamanoPagina INT = 6
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @PaginaFinal INT = CASE WHEN @Pagina IS NULL OR @Pagina < 1 THEN 1 ELSE @Pagina END;
        DECLARE @TamanoPaginaFinal INT = CASE WHEN @TamanoPagina IS NULL OR @TamanoPagina < 1 THEN 6 ELSE @TamanoPagina END;
        DECLARE @Offset INT = (@PaginaFinal - 1) * @TamanoPaginaFinal;

        SELECT
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
            ) AS SedeMapaUrl,
            COUNT(1) OVER() AS TotalRegistros,
            CAST(CASE
                WHEN r.Estado IN (2, 3, 4) AND rr.Id IS NULL THEN 1
                ELSE 0
            END AS BIT) AS PuedeRegistrarResena,
            rr.Id AS ResenaId,
            rr.AliasPublico AS ResenaAliasPublico,
            rr.Comentario AS ResenaComentario,
            rr.FechaCreacion AS ResenaFechaCreacion,
            rr.Activo AS ResenaActivo,
            rr.Respuesta AS ResenaRespuesta
        FROM dbo.ReservasUsuariosPublicos rup
        INNER JOIN dbo.Reservas r ON r.Id = rup.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        LEFT JOIN dbo.ReservasUsuariosPublicosResenas rr ON rr.ReservaId = r.Id
        WHERE rup.UsuarioId = @UsuarioId
        ORDER BY r.Fecha DESC, r.HoraInicio DESC, r.Id DESC
        OFFSET @Offset ROWS FETCH NEXT @TamanoPaginaFinal ROWS ONLY;
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
