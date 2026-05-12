USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/05/2026
-- Description:   Obtiene contexto para correo de desafio recibido (destino, desafiante y datos del desafio).
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_ObtenerContextoEmail]
    @IdDesafio INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @IdDesafio IS NULL OR @IdDesafio <= 0
            RAISERROR('El id de desafio es obligatorio.', 16, 1);

        SELECT
            d.Id AS DesafioId,
            pRetado.Correo AS CorreoRetado,
            LTRIM(RTRIM(CONCAT(pRetado.Nombres, N' ', pRetado.Apellidos))) AS NombreRetado,
            COALESCE(NULLIF(LTRIM(RTRIM(pRetador.NombreEquipo)), N''), LTRIM(RTRIM(CONCAT(pRetador.Nombres, N' ', pRetador.Apellidos)))) AS EquipoRetador,
            LTRIM(RTRIM(CONCAT(pRetador.Nombres, N' ', pRetador.Apellidos))) AS ContactoRetador,
            uRetador.UserName AS UsuarioRetador,
            pRetador.Telefono AS TelefonoRetador,
            tsm.Nombre AS Deporte,
            nd.Nombre AS Nivel,
            CONCAT(ud.Nombre, N', ', up.Nombre, N', ', udp.Nombre) AS Distrito,
            d.FechaTentativa,
            d.HoraTentativa,
            d.CanchaSugerida,
            d.Modalidad,
            d.Mensaje,
            d.FormaPago
        FROM dbo.Desafio d
        INNER JOIN dbo.UsuariosPublicosPerfil pRetador
            ON pRetador.UsuarioId = d.IdUsuarioRetador
        INNER JOIN dbo.UsuariosPublicosPerfil pRetado
            ON pRetado.UsuarioId = d.IdUsuarioRetado
        INNER JOIN dbo.AspNetUsers uRetador
            ON uRetador.Id = d.IdUsuarioRetador
        INNER JOIN dbo.TiposDeporteSuperMaestro tsm
            ON tsm.Id = d.IdDeporte
           AND tsm.Activo = 1
        INNER JOIN dbo.NivelDesafio nd
            ON nd.IdNivel = d.IdNivel
        INNER JOIN dbo.UbigeoDistritos ud
            ON ud.CodigoUbigeo = d.Distrito
        INNER JOIN dbo.UbigeoProvincias up
            ON up.CodigoProvincia = ud.CodigoProvincia
        INNER JOIN dbo.UbigeoDepartamentos udp
            ON udp.CodigoDepartamento = ud.CodigoDepartamento
        WHERE d.Id = @IdDesafio
          AND d.Activo = 1;
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
