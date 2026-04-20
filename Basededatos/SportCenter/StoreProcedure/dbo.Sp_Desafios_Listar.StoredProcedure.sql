USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Lista desafios del usuario portal con detalle del rival, contacto condicionado por estado y soporte de coordinacion.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Description:   Devuelve ubicacion completa del desafio con distrito, provincia y departamento para la vista operacional.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Description:   Incluye el contacto responsable del equipo rival mediante nombre del perfil y usuario registrado.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_Listar]
    @UsuarioId NVARCHAR(450),
    @TipoListado NVARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @TipoListadoNorm NVARCHAR(20) = LOWER(LTRIM(RTRIM(@TipoListado)));

        SELECT
            d.Id,
            CASE
                WHEN d.IdUsuarioRetador = @UsuarioId THEN COALESCE(NULLIF(LTRIM(RTRIM(pRetado.NombreEquipo)), N''), LTRIM(RTRIM(CONCAT(pRetado.Nombres, N' ', pRetado.Apellidos))))
                ELSE COALESCE(NULLIF(LTRIM(RTRIM(pRetador.NombreEquipo)), N''), LTRIM(RTRIM(CONCAT(pRetador.Nombres, N' ', pRetador.Apellidos))))
            END AS RivalNombre,
            CASE
                WHEN d.IdUsuarioRetador = @UsuarioId THEN LTRIM(RTRIM(CONCAT(pRetado.Nombres, N' ', pRetado.Apellidos)))
                ELSE LTRIM(RTRIM(CONCAT(pRetador.Nombres, N' ', pRetador.Apellidos)))
            END AS ContactoNombreRival,
            CASE
                WHEN d.IdUsuarioRetador = @UsuarioId THEN uRetado.UserName
                ELSE uRetador.UserName
            END AS ContactoUsuarioRival,
            CASE WHEN d.IdUsuarioRetador = @UsuarioId THEN N'Retador' ELSE N'Retado' END AS RolVista,
            td.Nombre AS Deporte,
            nd.Nombre AS Nivel,
            CONCAT(ud.Nombre, N', ', up.Nombre, N', ', udp.Nombre) AS Distrito,
            d.FechaTentativa,
            d.HoraTentativa,
            d.CanchaSugerida,
            d.Modalidad,
            d.Mensaje,
            d.FormaPago,
            d.Estado,
            d.FechaCreacion,
            d.FechaRespuesta,
            CASE
                WHEN d.IdUsuarioRetador = @UsuarioId THEN pRetado.ObservacionDesafio
                ELSE pRetador.ObservacionDesafio
            END AS ObservacionDesafioRival,
            CASE
                WHEN d.IdUsuarioRetador = @UsuarioId THEN pRetado.DetalleEquipo
                ELSE pRetador.DetalleEquipo
            END AS DetalleEquipoRival,
            CASE
                WHEN d.Estado IN (N'Aceptado', N'Finalizado') AND d.IdUsuarioRetador = @UsuarioId THEN pRetado.Telefono
                WHEN d.Estado IN (N'Aceptado', N'Finalizado') THEN pRetador.Telefono
                ELSE NULL
            END AS TelefonoRival,
            CASE
                WHEN d.Estado IN (N'Aceptado', N'Finalizado') AND d.IdUsuarioRetador = @UsuarioId THEN pRetado.WhatsappEquipo
                WHEN d.Estado IN (N'Aceptado', N'Finalizado') THEN pRetador.WhatsappEquipo
                ELSE NULL
            END AS WhatsappRival,
            CASE WHEN d.Estado IN (N'Aceptado', N'Finalizado') THEN CAST(1 AS BIT) ELSE CAST(0 AS BIT) END AS PuedeVerContactoRival
        FROM dbo.Desafio d
        INNER JOIN dbo.TiposDeporte td
            ON td.Id = d.IdDeporte
        INNER JOIN dbo.NivelDesafio nd
            ON nd.IdNivel = d.IdNivel
        INNER JOIN dbo.UbigeoDistritos ud
            ON ud.CodigoUbigeo = d.Distrito
        INNER JOIN dbo.UbigeoProvincias up
            ON up.CodigoProvincia = ud.CodigoProvincia
        INNER JOIN dbo.UbigeoDepartamentos udp
            ON udp.CodigoDepartamento = ud.CodigoDepartamento
        INNER JOIN dbo.UsuariosPublicosPerfil pRetador
            ON pRetador.UsuarioId = d.IdUsuarioRetador
        INNER JOIN dbo.UsuariosPublicosPerfil pRetado
            ON pRetado.UsuarioId = d.IdUsuarioRetado
        INNER JOIN dbo.AspNetUsers uRetador
            ON uRetador.Id = d.IdUsuarioRetador
        INNER JOIN dbo.AspNetUsers uRetado
            ON uRetado.Id = d.IdUsuarioRetado
        WHERE d.Activo = 1
          AND (
                (@TipoListadoNorm = N'enviados' AND d.IdUsuarioRetador = @UsuarioId AND d.Estado <> N'Finalizado')
             OR (@TipoListadoNorm = N'recibidos' AND d.IdUsuarioRetado = @UsuarioId AND d.Estado <> N'Finalizado')
             OR (@TipoListadoNorm = N'historial' AND (d.IdUsuarioRetador = @UsuarioId OR d.IdUsuarioRetado = @UsuarioId) AND d.Estado = N'Finalizado')
          )
        ORDER BY d.FechaCreacion DESC, d.Id DESC;
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
