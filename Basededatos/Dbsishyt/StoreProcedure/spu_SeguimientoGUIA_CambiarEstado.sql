-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Cambia estado con dependencias, fecha, responsable, auditoría e idempotencia.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Guía como documento sin estados; validaciones por existencia.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Comprueba el permiso específico del perfil antes de realizar la acción.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'GUIAS.MODIFICAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'GUIAS.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        THROW 51431,N'[HYT_VALIDACION] La guía no utiliza estados. Modifica los datos desde su registro.',1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
