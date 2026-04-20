USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Obtiene perfil del usuario publico por UsuarioId.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Amplia el perfil publico con configuracion de desafios e informacion general del equipo.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Incorpora ubigeo propio y WhatsApp del equipo para la busqueda, visualizacion y coordinacion de desafios.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_UsuariosPublicos_ObtenerPerfil]
    @UsuarioId NVARCHAR(450)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            p.Id,
            p.UsuarioId,
            p.TipoDocumento,
            p.NumeroDocumento,
            p.Nombres,
            p.Apellidos,
            p.NombreEquipo,
            p.Telefono,
            p.Correo,
            p.FechaNacimiento,
            p.CodigoUbigeo,
            CASE WHEN p.CodigoUbigeo IS NOT NULL THEN LEFT(p.CodigoUbigeo, 2) END AS CodigoDepartamento,
            CASE WHEN p.CodigoUbigeo IS NOT NULL THEN LEFT(p.CodigoUbigeo, 4) END AS CodigoProvincia,
            p.BuscarDesafios,
            p.IdDeporteDesafio,
            p.IdNivelDesafio,
            p.ObservacionDesafio,
            p.DetalleEquipo,
            p.CodigoUbigeoEquipo,
            CASE WHEN p.CodigoUbigeoEquipo IS NOT NULL THEN LEFT(p.CodigoUbigeoEquipo, 2) END AS CodigoDepartamentoEquipo,
            CASE WHEN p.CodigoUbigeoEquipo IS NOT NULL THEN LEFT(p.CodigoUbigeoEquipo, 4) END AS CodigoProvinciaEquipo,
            p.WhatsappEquipo
        FROM dbo.UsuariosPublicosPerfil p
        WHERE p.UsuarioId = @UsuarioId;
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
