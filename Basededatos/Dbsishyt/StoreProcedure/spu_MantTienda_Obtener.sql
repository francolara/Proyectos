-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta Tienda, referencias e historial paginado.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos automáticos inalterables y retiro de posición del maestro de materiales.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTienda_Obtener
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
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.TiendaId AS Id,
            m.ClienteId,
            m.Codigo,
            m.Nombre,
            m.Direccion,
            m.CodigoUbigeo,
            m.CentroSAPFICO,
            m.ZonaId,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            CASE WHEN ref0.TipoEntidad=N'NATURAL' THEN CONCAT(ref0.Nombres, N' ', ref0.Apellidos) ELSE ref0.RazonSocial END AS ClienteIdLabel,
            ref6.Descripcion AS ZonaIdLabel
        FROM dbo.Tienda AS m LEFT JOIN dbo.Cliente AS ref0 ON ref0.ClienteId = m.ClienteId LEFT JOIN dbo.Zona AS ref6 ON ref6.Id = m.ZonaId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.TiendaId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Tienda' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Tienda' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
