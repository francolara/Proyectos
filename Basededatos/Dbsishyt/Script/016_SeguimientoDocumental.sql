-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Integridad y permisos de ejecución de trabajo, sin modificar presupuestos históricos.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.Ejecucion') AND name=N'UX_Ejecucion_Presupuesto')
BEGIN
    IF EXISTS (SELECT ejecucion.PresupuestoId FROM dbo.Ejecucion AS ejecucion GROUP BY ejecucion.PresupuestoId HAVING COUNT_BIG(1)>1)
        THROW 51401,N'Existen ejecuciones duplicadas por presupuesto. Revisar antes de crear la restricción.',1;
    CREATE UNIQUE INDEX UX_Ejecucion_Presupuesto ON dbo.Ejecucion(PresupuestoId);
END;
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_OC_ImporteNoNegativo')
    ALTER TABLE dbo.OrdenCompraCliente WITH CHECK ADD CONSTRAINT CK_OC_ImporteNoNegativo CHECK (Importe IS NULL OR Importe>=0);
GO
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    INSERT INTO dbo.Permiso (Codigo,Descripcion)
    SELECT origen.Codigo,origen.Descripcion FROM (VALUES (N'EJECUCION.CONSULTAR',N'Consultar ejecución del trabajo'),(N'EJECUCION.EDITAR',N'Registrar ejecución del trabajo')) AS origen(Codigo,Descripcion)
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK,HOLDLOCK) WHERE permiso.Codigo=origen.Codigo);
    INSERT INTO dbo.RolPermiso (RoleId,CodigoPermiso)
    SELECT roles.Id,permiso.Codigo FROM dbo.AspNetRoles AS roles CROSS JOIN dbo.Permiso AS permiso
    WHERE ((permiso.Codigo=N'EJECUCION.CONSULTAR' AND roles.Id IN (N'HYT-ADMIN',N'HYT-OPERACIONES',N'HYT-COMERCIAL',N'HYT-SEGUIMIENTO',N'HYT-CONSULTA'))
        OR (permiso.Codigo=N'EJECUCION.EDITAR' AND roles.Id IN (N'HYT-ADMIN',N'HYT-OPERACIONES')))
        AND NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS asignacion WITH (UPDLOCK,HOLDLOCK) WHERE asignacion.RoleId=roles.Id AND asignacion.CodigoPermiso=permiso.Codigo);
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
    SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
    RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
END CATCH;
GO
