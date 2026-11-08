-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Recupera HTML histórico conservado con permiso de presupuesto e impresión.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permiso de descarga de la impresión.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_ObtenerImpresion
    @UsuarioId NVARCHAR(128),@ImpresionId BIGINT,@Descargar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Descargar=1 AND dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.DESCARGAR')=0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para descargar el presupuesto.',1;
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar impresiones.',1;
        SELECT impresion.ImpresionId,impresion.RevisionId,revision.PresupuestoId,impresion.Formato,impresion.VersionFormato,impresion.FechaRegistro,
            usuario.Nombres AS Usuario,impresion.ContenidoHTML,impresion.SHA256
        FROM dbo.ImpresionPresupuesto AS impresion INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=impresion.RevisionId
        INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=impresion.UsuarioRegistro WHERE impresion.ImpresionId=@ImpresionId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
