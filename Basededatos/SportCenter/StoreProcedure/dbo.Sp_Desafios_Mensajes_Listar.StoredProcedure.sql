USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Lista mensajes de desafios para un participante autorizado en orden cronologico.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_Mensajes_Listar]
    @UsuarioId NVARCHAR(450),
    @IdDesafio INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            dm.IdMensaje,
            dm.IdDesafio,
            dm.UsuarioIdEmisor,
            COALESCE(NULLIF(LTRIM(RTRIM(pp.NombreEquipo)), N''), LTRIM(RTRIM(CONCAT(pp.Nombres, N' ', pp.Apellidos)))) AS NombreEmisor,
            CASE WHEN dm.UsuarioIdEmisor = @UsuarioId THEN CAST(1 AS BIT) ELSE CAST(0 AS BIT) END AS EsMio,
            dm.Mensaje,
            dm.FechaRegistro
        FROM dbo.DesafioMensaje dm
        INNER JOIN dbo.Desafio d
            ON d.Id = dm.IdDesafio
        INNER JOIN dbo.UsuariosPublicosPerfil pp
            ON pp.UsuarioId = dm.UsuarioIdEmisor
        WHERE dm.Activo = 1
          AND d.Activo = 1
          AND (d.IdUsuarioRetador = @UsuarioId OR d.IdUsuarioRetado = @UsuarioId)
          AND (@IdDesafio IS NULL OR dm.IdDesafio = @IdDesafio)
        ORDER BY dm.IdDesafio ASC, dm.FechaRegistro ASC, dm.IdMensaje ASC;
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
