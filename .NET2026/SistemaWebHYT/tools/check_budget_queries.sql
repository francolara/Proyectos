-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Prueba local de consultas de presupuesto con actor efímero y rollback.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    DECLARE @Actor NVARCHAR(128)=CONCAT(N'test-presupuesto-',CONVERT(NVARCHAR(36),NEWID()));
    DECLARE @Rol NVARCHAR(128);
    SELECT TOP (1) @Rol=permiso.RoleId FROM dbo.RolPermiso AS permiso WHERE permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' ORDER BY permiso.RoleId;
    IF @Rol IS NULL THROW 51400,N'Falta rol de consulta para la prueba.',1;
    INSERT INTO dbo.AspNetUsers (Id,Nombres,Activo) VALUES (@Actor,N'Prueba transaccional',1);
    INSERT INTO dbo.AspNetUserRoles (UserId,RoleId) VALUES (@Actor,@Rol);
    EXEC dbo.spu_Presupuesto_Listar @UsuarioId=@Actor,@Desde='2026-10-03',@HastaExclusiva='2026-11-04';
    EXEC dbo.spu_Presupuesto_Obtener @UsuarioId=@Actor,@Id=-1;
    EXEC dbo.spu_Presupuesto_AyudaTarifario @UsuarioId=@Actor,@ClienteId=-1,@Fecha='2026-11-03',@CodigoMoneda=N'USD';
    ROLLBACK TRANSACTION;
    PRINT N'Consultas de presupuesto verificadas; actor de prueba revertido.';
END TRY
BEGIN CATCH
    IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
    SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
    RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
END CATCH;
GO
