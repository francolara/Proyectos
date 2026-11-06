-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta Cliente, referencias e historial paginado.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Consulta identidad, Código SAP y enlaces de documentos ordenados con historial del cliente.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Obtener
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
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.ClienteId AS Id,
            m.Codigo,
            m.CodigoSAP,
            m.TipoEntidad,
            m.TipoIdentificacion,
            m.NumeroIdentificacion,
            m.RazonSocial,
            m.Nombres,
            m.Apellidos,
            m.DireccionFiscal,
            m.CodigoUbigeo,
            m.UrlCargaDocumentos,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=m.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces
        FROM dbo.Cliente AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.ClienteId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Cliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Cliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
