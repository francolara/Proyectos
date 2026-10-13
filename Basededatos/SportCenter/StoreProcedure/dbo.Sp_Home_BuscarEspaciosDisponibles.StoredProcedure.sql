
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
-- Firma: FRANCO LARA - 09/06/2026 | Limita el home publico a espacios de negocios con suscripcion publica habilitada (EstadoSuscripcion = 1 o 2), ocultando pendientes, vencidos y suspendidos en tarjetas y resultados directos.
-- Firma: FRANCO LARA - 06/10/2026 | Expone moneda canonica y resuelve deporte/suelo desde los identificadores globales persistidos por el espacio.
-- Firma: Codex - 13/10/2026 | Alinea la disponibilidad publica con reservas: valida horario efectivo de sede/espacio, dia habilitado, fechas inhabilitadas, bloqueos y cruces en espacios relacionados.
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
        ;WITH EspaciosAfectados AS
        (
            SELECT
                e.Id AS EspacioBaseId,
                e.Id AS EspacioAfectadoId
            FROM dbo.EspaciosDeportivos e
            WHERE e.Estado = 1

            UNION

            SELECT
                e.Id AS EspacioBaseId,
                ec.EspacioRelacionadoId AS EspacioAfectadoId
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.EspaciosDeportivosCompartidos ec
                ON ec.EspacioDeportivoId = e.Id
               AND ec.Activo = 1
               AND ec.TipoRelacion IN (N'DIRECTO', N'COMPUESTO_COMPONENTE')
            INNER JOIN dbo.EspaciosDeportivos er
                ON er.Id = ec.EspacioRelacionadoId
               AND er.Estado = 1

            UNION

            SELECT
                e.Id AS EspacioBaseId,
                ec.EspacioDeportivoId AS EspacioAfectadoId
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.EspaciosDeportivosCompartidos ec
                ON ec.EspacioRelacionadoId = e.Id
               AND ec.Activo = 1
               AND ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
            INNER JOIN dbo.EspaciosDeportivos er
                ON er.Id = ec.EspacioDeportivoId
               AND er.Estado = 1

            UNION

            SELECT
                e.Id AS EspacioBaseId,
                ed.EspacioRelacionadoId AS EspacioAfectadoId
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.EspaciosDeportivosCompartidos ecComp
                ON ecComp.EspacioDeportivoId = e.Id
               AND ecComp.Activo = 1
               AND ecComp.TipoRelacion = N'COMPUESTO_COMPONENTE'
            INNER JOIN dbo.EspaciosDeportivosCompartidos ed
                ON ed.EspacioDeportivoId = ecComp.EspacioRelacionadoId
               AND ed.Activo = 1
               AND ed.TipoRelacion = N'DIRECTO'
            INNER JOIN dbo.EspaciosDeportivos er
                ON er.Id = ed.EspacioRelacionadoId
               AND er.Estado = 1

            UNION

            SELECT
                e.Id AS EspacioBaseId,
                ep.EspacioDeportivoId AS EspacioAfectadoId
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.EspaciosDeportivosCompartidos edActual
                ON edActual.EspacioDeportivoId = e.Id
               AND edActual.Activo = 1
               AND edActual.TipoRelacion = N'DIRECTO'
            INNER JOIN dbo.EspaciosDeportivosCompartidos ep
                ON ep.EspacioRelacionadoId = edActual.EspacioRelacionadoId
               AND ep.Activo = 1
               AND ep.TipoRelacion = N'COMPUESTO_COMPONENTE'
            INNER JOIN dbo.EspaciosDeportivos er
                ON er.Id = ep.EspacioDeportivoId
               AND er.Estado = 1
        ),
        Resultados AS
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
                END AS DistanciaKm,
                tarifaMin.CodigoMoneda,
                tarifaMin.MonedaSimbolo
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
            LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
            INNER JOIN dbo.TiposDeporteSuperMaestro td ON td.Id = e.TipoDeporteSuperId
            LEFT JOIN dbo.TiposSueloSuperMaestro ts ON ts.Id = e.TipoSueloSuperId
            LEFT JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = s.CodigoUbigeo
            LEFT JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
            LEFT JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento
            LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
            LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
            LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
            OUTER APPLY
            (
                SELECT TOP (1)
                    t.Precio AS TarifaDesde,
                    t.CodigoMoneda,
                    COALESCE(ms.Simbolo, t.CodigoMoneda) AS MonedaSimbolo
                FROM dbo.Tarifas t
                LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Codigo = t.CodigoMoneda
                WHERE t.EspacioDeportivoId = e.Id
                  AND t.Activa = 1
                ORDER BY t.Precio, t.Id
            ) tarifaMin
            WHERE e.Estado = 1
              AND COALESCE(e.AdministracionPrivada, 0) = 0
              AND s.Activo = 1
              AND n.Activo = 1
              AND COALESCE(ns.EstadoSuscripcion, 0) IN (1, 2)
              AND (
                    @TipoDeporteId IS NULL
                    OR e.TipoDeporteSuperId = @TipoDeporteId
                  )
              AND (@NegocioId IS NULL OR n.Id = @NegocioId)
              AND (@CodigoDepartamento IS NULL OR (s.CodigoUbigeo IS NOT NULL AND LEFT(s.CodigoUbigeo, 2) = @CodigoDepartamento))
              AND (@CodigoProvincia IS NULL OR (s.CodigoUbigeo IS NOT NULL AND LEFT(s.CodigoUbigeo, 4) = @CodigoProvincia))
              AND (@CodigoUbigeo IS NULL OR s.CodigoUbigeo = @CodigoUbigeo)
              AND
              (
                  @IgnorarFechaHorario = 1
                  OR
                  (
                      @HoraFin > @HoraInicio
                      AND CASE (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1
                              WHEN 1 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeLunes, 1) ELSE COALESCE(sha.AtiendeLunes, 1) END
                              WHEN 2 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeMartes, 1) ELSE COALESCE(sha.AtiendeMartes, 1) END
                              WHEN 3 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeMiercoles, 1) ELSE COALESCE(sha.AtiendeMiercoles, 1) END
                              WHEN 4 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeJueves, 1) ELSE COALESCE(sha.AtiendeJueves, 1) END
                              WHEN 5 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeViernes, 1) ELSE COALESCE(sha.AtiendeViernes, 1) END
                              WHEN 6 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeSabado, 1) ELSE COALESCE(sha.AtiendeSabado, 1) END
                              WHEN 7 THEN CASE WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1 THEN COALESCE(eha.AtiendeDomingo, 1) ELSE COALESCE(sha.AtiendeDomingo, 1) END
                              ELSE 0
                          END = 1
                      AND @HoraInicio >= CASE
                              WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1
                                  THEN COALESCE(eha.HoraApertura, CAST('08:00' AS TIME))
                              ELSE COALESCE(sha.HoraApertura, CAST('08:00' AS TIME))
                          END
                      AND @HoraFin <= CASE
                              WHEN COALESCE(eha.ConfigurarHorarioPorEspacio, 0) = 1
                                  THEN COALESCE(eha.HoraCierre, CAST('23:00' AS TIME))
                              ELSE COALESCE(sha.HoraCierre, CAST('23:00' AS TIME))
                          END
                      AND NOT EXISTS
                      (
                          SELECT 1
                          FROM dbo.SedeFechasInhabilitadas sfi
                          WHERE sfi.SedeId = s.Id
                            AND sfi.Fecha = @Fecha
                            AND sfi.Activo = 1
                      )
                      AND NOT EXISTS
                      (
                          SELECT 1
                          FROM dbo.Reservas r
                          INNER JOIN EspaciosAfectados ea ON ea.EspacioAfectadoId = r.EspacioDeportivoId
                          WHERE ea.EspacioBaseId = e.Id
                            AND r.Fecha = @Fecha
                            AND r.Estado NOT IN (5, 6)
                            AND @HoraInicio < r.HoraFin
                            AND @HoraFin > r.HoraInicio
                      )
                      AND NOT EXISTS
                      (
                          SELECT 1
                          FROM dbo.BloqueosHorario b
                          INNER JOIN EspaciosAfectados ea ON ea.EspacioAfectadoId = b.EspacioDeportivoId
                          WHERE ea.EspacioBaseId = e.Id
                            AND b.Fecha = @Fecha
                            AND b.Activo = 1
                            AND @HoraInicio < b.HoraFin
                            AND @HoraFin > b.HoraInicio
                      )
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
                END AS DistanciaKm,
                CAST(N'PEN' AS NVARCHAR(10)) AS CodigoMoneda,
                CAST(N'S/' AS NVARCHAR(10)) AS MonedaSimbolo
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
            DistanciaKm,
            CodigoMoneda,
            MonedaSimbolo
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
            DistanciaKm,
            CodigoMoneda,
            MonedaSimbolo
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
                DistanciaKm,
                CodigoMoneda,
                MonedaSimbolo
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
                DistanciaKm,
                CodigoMoneda,
                MonedaSimbolo
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
