USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Busca rivales disponibles para desafios filtrando obligatoriamente por distrito.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   La busqueda compara el distrito solicitado contra la ubicacion del equipo, no la personal.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Description:   Expone el contacto responsable del equipo usando nombre del perfil y usuario registrado.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/05/2026
-- Description:   Usa TiposDeporteSuperMaestro como fuente de deporte publico para evitar dependencias con catalogo por negocio.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_BuscarRivales]
    @UsuarioId NVARCHAR(450),
    @CodigoUbigeo CHAR(6),
    @IdDeporte INT = NULL,
    @IdNivel INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @CodigoUbigeoNorm CHAR(6) = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');

        IF @CodigoUbigeoNorm IS NULL
            RAISERROR('Debes seleccionar un distrito para buscar rivales.', 16, 1);

        SELECT
            p.Id,
            p.UsuarioId,
            COALESCE(NULLIF(LTRIM(RTRIM(p.NombreEquipo)), N''), LTRIM(RTRIM(CONCAT(p.Nombres, N' ', p.Apellidos)))) AS NombreEquipo,
            LTRIM(RTRIM(CONCAT(p.Nombres, N' ', p.Apellidos))) AS ContactoNombre,
            u.UserName AS ContactoUsuario,
            ud.Nombre AS Distrito,
            tsm.Nombre AS Deporte,
            nd.Nombre AS Nivel,
            p.ObservacionDesafio,
            p.DetalleEquipo,
            p.IdDeporteDesafio,
            p.IdNivelDesafio,
            p.CodigoUbigeoEquipo,
            p.BuscarDesafios
        FROM dbo.UsuariosPublicosPerfil p
        INNER JOIN dbo.AspNetUsers u
            ON u.Id = p.UsuarioId
        INNER JOIN dbo.UbigeoDistritos ud
            ON ud.CodigoUbigeo = p.CodigoUbigeoEquipo
        INNER JOIN dbo.TiposDeporteSuperMaestro tsm
            ON tsm.Id = p.IdDeporteDesafio
           AND tsm.Activo = 1
        INNER JOIN dbo.NivelDesafio nd
            ON nd.IdNivel = p.IdNivelDesafio
        WHERE p.BuscarDesafios = 1
          AND p.UsuarioId <> @UsuarioId
          AND p.CodigoUbigeoEquipo = @CodigoUbigeoNorm
          AND (@IdDeporte IS NULL OR p.IdDeporteDesafio = @IdDeporte)
          AND (@IdNivel IS NULL OR p.IdNivelDesafio = @IdNivel)
        ORDER BY NombreEquipo, tsm.Nombre, nd.Orden;
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
