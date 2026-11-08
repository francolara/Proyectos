-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Datos únicos HYT y auditoría paginada para configuración corporativa.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo en direcciones y consideraciones predeterminadas por presupuesto.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Ubigeo HYT y metadatos de firma; lectura corporativa para impresiones autorizadas. Subtítulo de cotización editable con auditoría.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConfiguracionHYT_Obtener
    @UsuarioId NVARCHAR(128),@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIGURACION.CONSULTAR')
            IF NOT (EXISTS(SELECT 1 FROM dbo.AspNetUsers AS u INNER JOIN dbo.AspNetUserRoles AS r ON r.UserId=u.Id INNER JOIN dbo.RolPermiso AS p ON p.RoleId=r.RoleId WHERE u.Id=@UsuarioId AND u.Activo=1 AND p.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
                AND EXISTS(SELECT 1 FROM dbo.AspNetUsers AS u INNER JOIN dbo.AspNetUserRoles AS r ON r.UserId=u.Id INNER JOIN dbo.RolPermiso AS p ON p.RoleId=r.RoleId WHERE u.Id=@UsuarioId AND u.Activo=1 AND p.CodigoPermiso=N'REPORTES.CONSULTAR'))
                THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar la configuración.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
        SELECT (SELECT empresa.Id,empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.CodigoUbigeo,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.SubtituloCotizacion,empresa.ConsideracionesPresupuesto,empresa.FirmaPresupuestoJson,
            empresa.FechaRegistro,creador.Nombres AS UsuarioRegistro,empresa.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,empresa.VersionFila
            FROM dbo.ConfiguracionHYT AS empresa INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=empresa.UsuarioRegistro
            LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=empresa.UsuarioActualizacion WHERE empresa.Id=1
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion,usuario.Nombres AS Usuario,evento.FechaRegistro,evento.EstadoAnterior,evento.EstadoNuevo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ConfiguracionHYT' AND evento.RegistroId=N'1'
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'ConfiguracionHYT' AND evento.RegistroId=N'1';
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
