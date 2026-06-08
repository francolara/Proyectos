
GO
/****** Object:  StoredProcedure [dbo].[Sp_Home_BuscarEspaciosDisponibles]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 20_Home_Espacios_Whatsapp.sql (linea 7)
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Incluye ConsideracionesReserva de la sede en resultados de espacios disponibles.
-- =============================================
-- Firma: Codex - 14/04/2026 | Filtra disponibilidad publica por departamento/provincia/distrito/negocio y enriquece tarjetas con ubicacion, tipo de suelo, tarifa, contacto de sede (correo/whatsapp) y fotos para mini carrusel.
-- Firma: Codex - 15/04/2026 | Agrega @IgnorarFechaHorario (solo para busqueda por negocio): permite listar todos los espacios del club sin filtrar cruce por fecha/hora cuando el usuario marca "obviar dia y horario" en Home.
-- Firma: Codex - 18/04/2026 | Excluye espacios con AdministracionPrivada=1 para que no aparezcan en el portal publico.
-- Firma: Codex - 27/04/2026 | Cambia filtros de ubigeo en Home para usar CodigoUbigeo de Sede (no Negocio), alinea deporte con TipoDeporteSuperId, agrega union con referenciales externos, expone GoogleMapsUrl/Telefono por fila y agrega busqueda "cerca de mi" por lat/long (sedes + externos) con radio en km.
-- Firma: Codex - 29/04/2026 | Agrega paginacion SQL real para Home con @Pagina/@TamanoPagina y salida @TotalRegistros para evitar paginacion en memoria; en referenciales externos retorna Codigo vacio para no exponer identificadores tecnicos en tarjetas publicas.
-- Firma: FRANCO LARA - 08/06/2026 | Expone fotos propias del espacio deportivo para priorizarlas en las tarjetas del Home, completa con fotos de la sede cuando existan y propaga las nuevas columnas en la tabla temporal de paginacion.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_BuscarEspaciosDisponibles]
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @TipoDeporteId INT = NULL,
    @NegocioId INT = NULL,
    @IgnorarFechaHorario BIT = 0,
    @BuscarCercaDeMi BIT = 0,
    @LatitudUsuario DECIMAL(10,7) = NULL,
    @LongitudUsuario DECIMAL(10,7) = NULL,
    @RadioKm DECIMAL(6,2) = NULL,
    @Pagina INT = NULL,
    @TamanoPagina INT = NULL,
    @TotalRegistros INT = NULL OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH Resultados AS
        (
            SELECT
                0 AS OrdenFuente,
                e.Id,
                e.Nombre,
                e.Codigo,
                s.Nombre AS SedeNombre,
                s.Direccion AS SedeDireccion,
                s.ConsideracionesReserva AS SedeConsideracionesReserva,
                dep.Nombre AS Departamento,
                prov.Nombre AS Provincia,
                dist.Nombre AS Distrito,
                td.Nombre AS TipoDeporte,
                ts.Nombre AS TipoSuelo,
                tarifaMin.TarifaDesde,
                e.TieneIluminacion,
                e.Techada,
                scn.CorreoNotificacion,
                s.Telefono AS TelefonoContacto,
                scn.WhatsappContacto,
                COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
                s.Id AS SedeId,
                s.GoogleMapsUrl AS SedeMapaUrl,
                e.FotoPrincipalUrl AS EspacioFotoPrincipalUrl,
                e.FotosUrlsCsv AS EspacioFotosUrlsCsv,
                s.FotoPrincipalUrl AS SedeFotoPrincipalUrl,
                s.FotosUrlsCsv AS SedeFotosUrlsCsv,
                CASE
                    WHEN @BuscarCercaDeMi = 1
                         AND @LatitudUsuario IS NOT NULL
                         AND @LongitudUsuario IS NOT NULL
                         AND s.Latitud IS NOT NULL
                         AND s.Longitud IS NOT NULL
                    THEN CAST(
                            geography::Point(CONVERT(float, @LatitudUsuario), CONVERT(float, @LongitudUsuario), 4326)
                                .STDistance(geography::Point(CONVERT(float, s.Latitud), CONVERT(float, s.Longitud), 4326)) / 1000.0
                         AS DECIMAL(10,2))
                    ELSE NULL
                END AS DistanciaKm
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
            INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
            LEFT JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
            LEFT JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = s.CodigoUbigeo
            LEFT JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
            LEFT JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento
            LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
            OUTER APPLY
            (
                SELECT MIN(t.Precio) AS TarifaDesde
                FROM dbo.Tarifas t
                WHERE t.EspacioDeportivoId = e.Id
                  AND t.Activa = 1
            ) tarifaMin
            WHERE e.Estado = 1
              AND COALESCE(e.AdministracionPrivada, 0) = 0
              AND s.Activo = 1
              AND n.Activo = 1
              AND (
                    @TipoDeporteId IS NULL
                    OR td.TipoDeporteSuperId = @TipoDeporteId
                  )
              AND (@NegocioId IS NULL OR n.Id = @NegocioId)
              AND (@CodigoDepartamento IS NULL OR (s.CodigoUbigeo IS NOT NULL AND LEFT(s.CodigoUbigeo, 2) = @CodigoDepartamento))
              AND (@CodigoProvincia IS NULL OR (s.CodigoUbigeo IS NOT NULL AND LEFT(s.CodigoUbigeo, 4) = @CodigoProvincia))
              AND (@CodigoUbigeo IS NULL OR s.CodigoUbigeo = @CodigoUbigeo)
              AND
              (
                  @IgnorarFechaHorario = 1
                  OR NOT EXISTS
                  (
                      SELECT 1
                      FROM dbo.Reservas r
                      WHERE r.EspacioDeportivoId = e.Id
                        AND r.Fecha = @Fecha
                        AND r.Estado NOT IN (5, 6)
                        AND @HoraInicio < r.HoraFin
                        AND @HoraFin > r.HoraInicio
                  )
              )

            UNION ALL

            SELECT
                1 AS OrdenFuente,
                (-1 * he.Id) AS Id,
                COALESCE(NULLIF(LTRIM(RTRIM(he.NombreEspacio)), ''), he.NombreComplejo) AS Nombre,
                CAST('' AS NVARCHAR(80)) AS Codigo,
                '-' AS SedeNombre,
                he.Direccion AS SedeDireccion,
                he.Referencia AS SedeConsideracionesReserva,
                depEx.Nombre AS Departamento,
                provEx.Nombre AS Provincia,
                distEx.Nombre AS Distrito,
                tsm.Nombre AS TipoDeporte,
                NULL AS TipoSuelo,
                he.TarifaReferencial AS TarifaDesde,
                COALESCE(he.TieneIluminacion, 0) AS TieneIluminacion,
                COALESCE(he.Techada, 0) AS Techada,
                he.CorreoContacto AS CorreoNotificacion,
                he.TelefonoContacto,
                he.WhatsappContacto,
                COALESCE(he.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
                NULL AS SedeId,
                he.GoogleMapsUrl AS SedeMapaUrl,
                NULL AS EspacioFotoPrincipalUrl,
                NULL AS EspacioFotosUrlsCsv,
                he.FotoPrincipalUrl AS SedeFotoPrincipalUrl,
                he.FotosUrlsCsv AS SedeFotosUrlsCsv,
                CASE
                    WHEN @BuscarCercaDeMi = 1
                         AND @LatitudUsuario IS NOT NULL
                         AND @LongitudUsuario IS NOT NULL
                         AND he.LatitudReferencia IS NOT NULL
                         AND he.LongitudReferencia IS NOT NULL
                    THEN CAST(
                            geography::Point(CONVERT(float, @LatitudUsuario), CONVERT(float, @LongitudUsuario), 4326)
                                .STDistance(geography::Point(CONVERT(float, he.LatitudReferencia), CONVERT(float, he.LongitudReferencia), 4326)) / 1000.0
                         AS DECIMAL(10,2))
                    ELSE NULL
                END AS DistanciaKm
            FROM dbo.HomeEspaciosReferencialesExternos he
            INNER JOIN dbo.TiposDeporteSuperMaestro tsm ON tsm.Id = he.TipoDeporteSuperId
            LEFT JOIN dbo.UbigeoDistritos distEx ON distEx.CodigoUbigeo = he.CodigoUbigeo
            LEFT JOIN dbo.UbigeoProvincias provEx ON provEx.CodigoProvincia = distEx.CodigoProvincia
            LEFT JOIN dbo.UbigeoDepartamentos depEx ON depEx.CodigoDepartamento = distEx.CodigoDepartamento
            WHERE he.Activo = 1
              AND tsm.Activo = 1
              AND @NegocioId IS NULL
              AND (@TipoDeporteId IS NULL OR he.TipoDeporteSuperId = @TipoDeporteId)
              AND (@CodigoDepartamento IS NULL OR LEFT(he.CodigoUbigeo, 2) = @CodigoDepartamento)
              AND (@CodigoProvincia IS NULL OR LEFT(he.CodigoUbigeo, 4) = @CodigoProvincia)
              AND (@CodigoUbigeo IS NULL OR he.CodigoUbigeo = @CodigoUbigeo)
        )
        , Filtrados AS
        (
            SELECT
            OrdenFuente,
            Id,
            Nombre,
            Codigo,
            SedeNombre,
            SedeDireccion,
            SedeConsideracionesReserva,
            Departamento,
            Provincia,
            Distrito,
            TipoDeporte,
            TipoSuelo,
            TarifaDesde,
            TieneIluminacion,
            Techada,
            CorreoNotificacion,
            TelefonoContacto,
            WhatsappContacto,
            PermiteChatWhatsapp,
            SedeId,
            SedeMapaUrl,
            EspacioFotoPrincipalUrl,
            EspacioFotosUrlsCsv,
            SedeFotoPrincipalUrl,
            SedeFotosUrlsCsv,
            DistanciaKm
            FROM Resultados
            WHERE
                @BuscarCercaDeMi = 0
                OR
                (
                    DistanciaKm IS NOT NULL
                    AND DistanciaKm <= COALESCE(NULLIF(@RadioKm, 0), 5)
                )
        )
        SELECT
            OrdenFuente,
            Id,
            Nombre,
            Codigo,
            SedeNombre,
            SedeDireccion,
            SedeConsideracionesReserva,
            Departamento,
            Provincia,
            Distrito,
            TipoDeporte,
            TipoSuelo,
            TarifaDesde,
            TieneIluminacion,
            Techada,
            CorreoNotificacion,
            TelefonoContacto,
            WhatsappContacto,
            PermiteChatWhatsapp,
            SedeId,
            SedeMapaUrl,
            EspacioFotoPrincipalUrl,
            EspacioFotosUrlsCsv,
            SedeFotoPrincipalUrl,
            SedeFotosUrlsCsv,
            DistanciaKm
        INTO #Filtrados
        FROM Filtrados;

        SELECT @TotalRegistros = COUNT(1)
        FROM #Filtrados;

        IF @Pagina IS NOT NULL AND @TamanoPagina IS NOT NULL AND @TamanoPagina > 0
        BEGIN
            SELECT
                Id,
                Nombre,
                Codigo,
                SedeNombre,
                SedeDireccion,
                SedeConsideracionesReserva,
                Departamento,
                Provincia,
                Distrito,
                TipoDeporte,
                TipoSuelo,
                TarifaDesde,
                TieneIluminacion,
                Techada,
                CorreoNotificacion,
                TelefonoContacto,
                WhatsappContacto,
                PermiteChatWhatsapp,
                SedeId,
                SedeMapaUrl,
                EspacioFotoPrincipalUrl,
                EspacioFotosUrlsCsv,
                SedeFotoPrincipalUrl,
                SedeFotosUrlsCsv,
                DistanciaKm
            FROM #Filtrados
            ORDER BY
                CASE WHEN @BuscarCercaDeMi = 1 THEN DistanciaKm ELSE NULL END,
                CASE WHEN @BuscarCercaDeMi = 0 THEN OrdenFuente ELSE 0 END,
                CASE WHEN @BuscarCercaDeMi = 1 THEN OrdenFuente ELSE 0 END,
                SedeNombre,
                Nombre
            OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
            FETCH NEXT @TamanoPagina ROWS ONLY;
        END
        ELSE
        BEGIN
            SELECT
                Id,
                Nombre,
                Codigo,
                SedeNombre,
                SedeDireccion,
                SedeConsideracionesReserva,
                Departamento,
                Provincia,
                Distrito,
                TipoDeporte,
                TipoSuelo,
                TarifaDesde,
                TieneIluminacion,
                Techada,
                CorreoNotificacion,
                TelefonoContacto,
                WhatsappContacto,
                PermiteChatWhatsapp,
                SedeId,
                SedeMapaUrl,
                EspacioFotoPrincipalUrl,
                EspacioFotosUrlsCsv,
                SedeFotoPrincipalUrl,
                SedeFotosUrlsCsv,
                DistanciaKm
            FROM #Filtrados
            ORDER BY
                CASE WHEN @BuscarCercaDeMi = 1 THEN DistanciaKm ELSE NULL END,
                CASE WHEN @BuscarCercaDeMi = 0 THEN OrdenFuente ELSE 0 END,
                CASE WHEN @BuscarCercaDeMi = 1 THEN OrdenFuente ELSE 0 END,
                SedeNombre,
                Nombre;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
