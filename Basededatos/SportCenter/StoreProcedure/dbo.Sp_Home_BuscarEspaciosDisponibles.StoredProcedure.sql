USE [DbSportCenter]
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
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_BuscarEspaciosDisponibles]
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @TipoDeporteId INT = NULL,
    @NegocioId INT = NULL,
    @IgnorarFechaHorario BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
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
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            s.Id AS SedeId,
            s.FotoPrincipalUrl AS SedeFotoPrincipalUrl,
            s.FotosUrlsCsv AS SedeFotosUrlsCsv
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        LEFT JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        LEFT JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = n.CodigoUbigeo
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
          AND s.Activo = 1
          AND n.Activo = 1
          AND (@TipoDeporteId IS NULL OR e.TipoDeporteId = @TipoDeporteId)
          AND (@NegocioId IS NULL OR n.Id = @NegocioId)
          AND (@CodigoDepartamento IS NULL OR (n.CodigoUbigeo IS NOT NULL AND LEFT(n.CodigoUbigeo, 2) = @CodigoDepartamento))
          AND (@CodigoProvincia IS NULL OR (n.CodigoUbigeo IS NOT NULL AND LEFT(n.CodigoUbigeo, 4) = @CodigoProvincia))
          AND (@CodigoUbigeo IS NULL OR n.CodigoUbigeo = @CodigoUbigeo)
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
        ORDER BY dep.Nombre, prov.Nombre, dist.Nombre, s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
