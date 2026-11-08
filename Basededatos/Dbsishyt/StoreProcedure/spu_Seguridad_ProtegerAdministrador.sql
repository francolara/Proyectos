-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Seguridad_ProtegerAdministrador; seguridad, concurrencia y auditoría según operación.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Exige conservar una cuenta habilitada con acceso y las acciones de administración esenciales.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguridad_ProtegerAdministrador
    
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            WHERE usuario.Activo=1 AND NULLIF(usuario.PasswordHash,N'') IS NOT NULL
                AND NULLIF(usuario.SecurityStamp,N'') IS NOT NULL
                AND NOT EXISTS
                (
                    SELECT requerido.Codigo
                    FROM (VALUES (N'USUARIOS.CONSULTAR'),(N'USUARIOS.CREAR'),(N'USUARIOS.MODIFICAR'),
                        (N'USUARIOS.ACTIVAR'),(N'USUARIOS.RESTABLECER'),(N'PERFILES.CONSULTAR'),(N'PERFILES.EDITAR')) AS requerido(Codigo)
                    WHERE NOT EXISTS
                    (
                        SELECT 1 FROM dbo.AspNetUserRoles AS asignacion
                        INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=asignacion.RoleId
                        WHERE asignacion.UserId=usuario.Id AND permiso.CodigoPermiso=requerido.Codigo
                    )
                )
        )
            THROW 51702,N'[HYT_VALIDACION] Debe quedar al menos un administrador activo con permisos para administrar usuarios y perfiles.',1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
