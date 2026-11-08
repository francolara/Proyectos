-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/11/2026
-- Description:   Consulta paginada y detalle de auditoría.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Auditoria_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'AUDITORIA.CONSULTAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar auditoría.',1;

        SELECT (SELECT evento.EventoId, evento.FechaRegistro, evento.Entidad, evento.Accion, evento.RegistroId,
            referencia.Referencia, CONCAT_WS(N' ',actor.Nombres,NULLIF(actor.Apellidos,N'')) AS Usuario, evento.FechaEvento, evento.Motivo,
            evento.EstadoAnterior, evento.EstadoNuevo, JSON_QUERY(evento.CambiosJson) AS Cambios
        FROM dbo.EventoAuditoria AS evento
        INNER JOIN dbo.AspNetUsers AS actor ON actor.Id=evento.UsuarioId
        LEFT JOIN dbo.AspNetUsers AS afectado ON evento.Entidad IN (N'Usuario', N'AspNetUsers', N'Acceso') AND afectado.Id=evento.RegistroId
        LEFT JOIN dbo.AspNetRoles AS perfil ON evento.Entidad IN (N'Perfil', N'AspNetRoles') AND perfil.Id=evento.RegistroId
        LEFT JOIN dbo.PresupuestoRevision AS revision ON evento.Entidad=N'PresupuestoRevision' AND revision.RevisionId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId=CASE
            WHEN evento.Entidad=N'Presupuesto' THEN TRY_CONVERT(BIGINT,evento.RegistroId)
            WHEN evento.Entidad=N'PresupuestoRevision' THEN revision.PresupuestoId END
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON evento.Entidad IN (N'OrdenCompraCliente',N'OC_CLIENTE') AND oc.OrdenCompraId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.GuiaConformidad AS guia ON evento.Entidad IN (N'GuiaConformidad',N'GUIA') AND guia.GuiaId=TRY_CONVERT(BIGINT,evento.RegistroId)
        CROSS APPLY (VALUES (COALESCE(afectado.Email,perfil.Name,presupuesto.NumeroPresupuesto,oc.NumeroOCCliente,guia.NumeroActa,evento.RegistroId))) AS referencia(Referencia)
        WHERE evento.EventoId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
