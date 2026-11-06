-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta TarifarioVersion, referencias e historial paginado.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.VersionId AS Id,
            m.TarifarioId,
            m.NumeroVersion,
            m.VigenciaDesde,
            m.VigenciaHasta,
            m.CodigoMonedaPredeterminada,
            m.CodigoTipoIGVPredeterminado,
            m.Estado,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            ref0.Nombre AS TarifarioIdLabel,
            JSON_QUERY((SELECT zona.ZonaId AS Id, catalogo.Codigo, catalogo.Descripcion FROM dbo.TarifarioVersionZona AS zona INNER JOIN dbo.Zona AS catalogo ON catalogo.Id=zona.ZonaId WHERE zona.VersionId = m.VersionId ORDER BY zona.ZonaId FOR JSON PATH)) AS Zonas
        FROM dbo.TarifarioVersion AS m INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId = m.TarifarioId LEFT JOIN dbo.Tarifario AS ref0 ON ref0.TarifarioId = m.TarifarioId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.VersionId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'TarifarioVersion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'TarifarioVersion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
