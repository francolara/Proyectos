-- Firma: FRANCO LARA | Create date: 17/11/2026 | Permite elegir una firma activa para reimprimir una revisión existente, conservando los datos del presupuesto.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_FirmaGuardar
    @UsuarioId NVARCHAR(128),@PresupuestoId BIGINT,@RevisionId BIGINT,@Version BINARY(8),@FirmaId BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.CONSULTAR')=0 OR dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.MODIFICAR')=0
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para elegir la firma.',1;
        BEGIN TRANSACTION;
        DECLARE @Archivo NVARCHAR(2000),@Nombre NVARCHAR(150),@Empresa NVARCHAR(MAX),@Anterior NVARCHAR(MAX);
        SELECT @Archivo=firma.ArchivoJson,@Nombre=firma.Nombre FROM dbo.FirmaPresupuestoHYT AS firma WITH(HOLDLOCK)
            WHERE firma.FirmaId=@FirmaId AND firma.Activo=1;
        IF @Archivo IS NULL THROW 51305,N'[HYT_VALIDACION] Selecciona una firma activa.',1;
        SELECT @Empresa=revision.EmpresaSnapshotJson FROM dbo.PresupuestoRevision AS revision WITH(UPDLOCK,HOLDLOCK)
            WHERE revision.PresupuestoId=@PresupuestoId AND revision.RevisionId=@RevisionId AND revision.VersionFila=@Version;
        IF @@ROWCOUNT=0 THROW 51309,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a abrirlo.',1;
        SET @Anterior=@Empresa;
        SET @Empresa=JSON_MODIFY(COALESCE(@Empresa,N'{}'),N'$.FirmaSeleccionada',CONVERT(BIT,1));
        SET @Empresa=JSON_MODIFY(@Empresa,N'$.FirmaId',@FirmaId);
        SET @Empresa=JSON_MODIFY(@Empresa,N'$.NombreFirma',@Nombre);
        SET @Empresa=JSON_MODIFY(@Empresa,N'$.FirmaPresupuestoJson',JSON_QUERY(@Archivo));
        UPDATE dbo.PresupuestoRevision SET EmpresaSnapshotJson=@Empresa,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
            WHERE RevisionId=@RevisionId;
        DECLARE @Operacion UNIQUEIDENTIFIER=NEWID(),@Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId);
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Empresa) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'PresupuestoRevision',@Registro,N'CAMBIAR_FIRMA',@CambiosJson=@Cambios;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_FirmaGuardar TO hyt_app_executor;
GO
