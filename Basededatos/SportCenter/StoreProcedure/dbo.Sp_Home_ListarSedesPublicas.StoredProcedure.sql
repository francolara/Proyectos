
GO
/****** Object:  StoredProcedure [dbo].[Sp_Home_ListarSedesPublicas]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 19_Home_Whatsapp_Publico.sql (linea 8)
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Incluye ConsideracionesReserva para mostrar condiciones de sede en portal publico.
-- =============================================
-- Firma: Codex - 14/04/2026 | Devuelve NegocioId y NombreComercial para poblar filtro publico por club/negocio.
-- Firma: Codex - 15/04/2026 | Agrega Servicios (catalogo de sede) para mostrar amenities del club en la vista publica de reserva.
-- Firma: Codex - 16/04/2026 | Expone URLs sociales por sede (Facebook/Instagram/Twitter) para iconos en tarjetas publicas y agrega codigos de ubigeo del negocio para filtrar combo de club en Home por departamento/provincia/distrito.
-- Firma: Codex - 27/04/2026 | Expone codigos de ubigeo por sede (manteniendo alias Negocio para compatibilidad), permitiendo filtrar club/negocio en Home segun ubicacion de sus sedes.
-- Firma: FRANCO LARA - 09/06/2026 | Limita el portal publico a sedes de negocios con suscripcion publica habilitada (EstadoSuscripcion = 1 o 2), ocultando pendientes, vencidos y suspendidos en el filtro de complejo deportivo.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ListarSedesPublicas]
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.Nombre,
            s.Direccion,
            s.ConsideracionesReserva,
            s.Telefono,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            s.FacebookUrl,
            s.InstagramUrl,
            s.TwitterUrl,
            s.Latitud,
            s.Longitud,
            s.GoogleMapsUrl,
            s.FotoPrincipalUrl,
            s.FotosUrlsCsv,
            n.Id AS NegocioId,
            n.NombreComercial AS NegocioNombre,
            STUFF((
                SELECT N', ' + cs.Nombre
                FROM dbo.SedeServicios ss
                INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = ss.ServicioId
                WHERE ss.SedeId = s.Id
                  AND cs.Activo = 1
                ORDER BY cs.Nombre
                FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)'), 1, 2, N'') AS Servicios,
            s.CodigoUbigeo AS CodigoUbigeoNegocio,
            CASE WHEN s.CodigoUbigeo IS NOT NULL THEN LEFT(s.CodigoUbigeo, 2) END AS CodigoDepartamentoNegocio,
            CASE WHEN s.CodigoUbigeo IS NOT NULL THEN LEFT(s.CodigoUbigeo, 4) END AS CodigoProvinciaNegocio
        FROM dbo.Sedes s
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE s.Activo = 1
          AND n.Activo = 1
          AND COALESCE(ns.EstadoSuscripcion, 0) IN (1, 2)
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
