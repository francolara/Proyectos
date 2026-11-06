-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista generaciones de impresión de una revisión con página de veinte.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_ListarImpresiones
    @UsuarioId NVARCHAR(128),@RevisionId BIGINT,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar impresiones.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
        SELECT impresion.ImpresionId,impresion.RevisionId,revision.PresupuestoId,impresion.Formato,impresion.VersionFormato,impresion.FechaRegistro,usuario.Nombres AS Usuario
        FROM dbo.ImpresionPresupuesto AS impresion INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=impresion.RevisionId
        INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=impresion.UsuarioRegistro WHERE impresion.RevisionId=@RevisionId
        ORDER BY impresion.ImpresionId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.ImpresionPresupuesto AS impresion WHERE impresion.RevisionId=@RevisionId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
