-- Firma: FRANCO LARA | Create date: 14/11/2026 | Consulta o reemplaza el adjunto SCTR con permisos, control de concurrencia y auditoría.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PolizaSctrHYT_Adjunto
    @UsuarioId NVARCHAR(128), @PolizaId BIGINT,
    @Nombre NVARCHAR(255)=NULL, @Clave NVARCHAR(500)=NULL, @Contenido NVARCHAR(150)=NULL,
    @Tamano BIGINT=NULL, @SHA256 CHAR(64)=NULL, @ArchivoAnterior BIGINT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'OBLIGACIONES.CONSULTAR')=0
            THROW 51000,N'No tiene permiso para consultar la póliza SCTR.',1;

        IF @Clave IS NULL
        BEGIN
            SELECT (SELECT archivo.NombreOriginal,archivo.ClaveAlmacenamiento,archivo.TipoContenido,
                archivo.TamanoBytes,archivo.SHA256
            FROM dbo.PolizaSctrHYT AS poliza
            INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=poliza.ArchivoId AND archivo.Retirado=0
            WHERE poliza.PolizaId=@PolizaId
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            RETURN;
        END;

        IF dbo.fn_HYTPermiso(@UsuarioId,N'OBLIGACIONES.EDITAR')=0
            THROW 51000,N'No tiene permiso para adjuntar la póliza SCTR.',1;
        DECLARE @Prefijo NVARCHAR(40)=CONCAT(N'SCTR/',@PolizaId,N'/');
        IF LEFT(@Clave,LEN(@Prefijo))<>@Prefijo OR @Nombre IS NULL OR @SHA256 IS NULL
            OR @Tamano IS NULL OR @Tamano NOT BETWEEN 1 AND 15728640
            OR @Contenido IS NULL OR @Contenido NOT IN (N'application/pdf',N'image/png',N'image/jpeg')
            THROW 51000,N'El adjunto de la póliza no es válido.',1;

        DECLARE @Actual BIGINT,@Existe BIGINT,@Nuevo BIGINT;
        BEGIN TRANSACTION;
        SELECT @Existe=poliza.PolizaId,@Actual=poliza.ArchivoId
        FROM dbo.PolizaSctrHYT AS poliza WITH(UPDLOCK,HOLDLOCK) WHERE poliza.PolizaId=@PolizaId;
        IF @Existe IS NULL THROW 51000,N'La póliza SCTR no existe.',1;
        IF EXISTS(SELECT 1 FROM dbo.ArchivoAdjunto AS archivo WHERE archivo.ArchivoId=@Actual
            AND archivo.ClaveAlmacenamiento=@Clave AND archivo.SHA256=@SHA256 AND archivo.Retirado=0)
        BEGIN
            COMMIT TRANSACTION;
            SELECT @Actual;
            RETURN;
        END;
        IF ISNULL(@Actual,0)<>ISNULL(@ArchivoAnterior,0)
            THROW 51000,N'El adjunto cambió. Vuelve a abrir la póliza.',1;

        INSERT dbo.ArchivoAdjunto(NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
        VALUES(@Nombre,@Clave,@Contenido,@Tamano,@SHA256,@UsuarioId);
        SET @Nuevo=SCOPE_IDENTITY();
        UPDATE dbo.PolizaSctrHYT SET ArchivoId=@Nuevo WHERE PolizaId=@PolizaId;
        UPDATE dbo.ArchivoAdjunto SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        WHERE ArchivoId=@Actual;
        DECLARE @Correlacion UNIQUEIDENTIFIER=NEWID(),@Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@PolizaId);
        DECLARE @Datos NVARCHAR(MAX)=(SELECT @Actual AS ArchivoAnterior,@Nuevo AS ArchivoNuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Correlacion,N'PolizaSctrHYT',@Registro,N'ADJUNTAR',@CambiosJson=@Datos;
        COMMIT TRANSACTION;
        SELECT @Nuevo;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_PolizaSctrHYT_Adjunto TO hyt_app_executor;
GO
