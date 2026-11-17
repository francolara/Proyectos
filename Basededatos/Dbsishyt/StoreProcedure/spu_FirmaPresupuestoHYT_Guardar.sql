-- Firma: FRANCO LARA | Create date: 17/11/2026 | Agrega una firma nombrada y permite establecer la predeterminada con auditoría.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_FirmaPresupuestoHYT_Guardar
    @UsuarioId NVARCHAR(128),@Nombre NVARCHAR(150)=NULL,@Datos NVARCHAR(2000)=NULL,
    @Predeterminada BIT=0,@FirmaId BIGINT=NULL,@Activo BIT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.CONSULTAR')=0 OR dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.EDITAR')=0
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para editar las firmas.',1;
        IF @FirmaId IS NULL
        BEGIN
            SET @Nombre=NULLIF(TRIM(@Nombre),N'');
            IF @Nombre IS NULL THROW 51305,N'[HYT_VALIDACION] Escribe un nombre para identificar la firma.',1;
        IF ISJSON(@Datos)<>1 OR @Datos IS NULL THROW 51303,N'[HYT_VALIDACION] Firma no válida.',1;
        DECLARE @Clave NVARCHAR(100)=JSON_VALUE(@Datos,N'$.ClaveAlmacenamiento'),@Tipo NVARCHAR(50)=JSON_VALUE(@Datos,N'$.TipoContenido'),@Hash NVARCHAR(100)=JSON_VALUE(@Datos,N'$.SHA256'),@Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TamanoBytes'));
        DECLARE @NombreClave NVARCHAR(100)=CASE WHEN LEFT(@Clave,16) COLLATE Latin1_General_100_BIN2=N'ConfigHYT/Firma/' THEN SUBSTRING(@Clave,17,100) ELSE @Clave END;
        IF @Clave IS NULL OR LEN(@NombreClave)<>36 OR LEFT(@NombreClave,32) COLLATE Latin1_General_100_BIN2 LIKE N'%[^a-f0-9]%' OR RIGHT(@Clave,4) NOT IN(N'.png',N'.jpg')
            OR @Tipo IS NULL OR @Tipo NOT IN(N'image/png',N'image/jpeg') OR (@Tipo=N'image/png' AND RIGHT(@Clave,4)<>N'.png') OR (@Tipo=N'image/jpeg' AND RIGHT(@Clave,4)<>N'.jpg')
            OR @Hash IS NULL OR LEN(@Hash)<>64 OR @Hash COLLATE Latin1_General_100_BIN2 LIKE N'%[^A-Fa-f0-9]%' OR @Tamano IS NULL OR @Tamano NOT BETWEEN 1 AND 2097152
            OR NULLIF(JSON_VALUE(@Datos,N'$.NombreOriginal'),N'') IS NULL
            THROW 51305,N'[HYT_VALIDACION] Adjunta una firma JPG o PNG válida de hasta 2 MB.',1;
        END;
        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:FirmasPresupuesto',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
        IF @Bloqueo<0 THROW 51305,N'[HYT_VALIDACION] Vuelve a intentar guardar la firma.',1;
        IF @FirmaId IS NULL
        BEGIN
            IF EXISTS(SELECT 1 FROM dbo.FirmaPresupuestoHYT WITH(UPDLOCK,HOLDLOCK) WHERE Nombre=@Nombre)
                THROW 51305,N'[HYT_VALIDACION] Ya existe una firma con ese nombre.',1;
            IF NOT EXISTS(SELECT 1 FROM dbo.FirmaPresupuestoHYT) SET @Predeterminada=1;
            INSERT INTO dbo.FirmaPresupuestoHYT(Nombre,ArchivoJson,Predeterminada,UsuarioRegistro)
            VALUES(@Nombre,@Datos,0,@UsuarioId);
            SET @FirmaId=SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SELECT @Datos=firma.ArchivoJson FROM dbo.FirmaPresupuestoHYT AS firma WHERE firma.FirmaId=@FirmaId;
            IF @Datos IS NULL THROW 51305,N'[HYT_VALIDACION] La firma seleccionada ya no existe.',1;
            IF @Activo IS NULL
            BEGIN
                IF NOT EXISTS(SELECT 1 FROM dbo.FirmaPresupuestoHYT WHERE FirmaId=@FirmaId AND Activo=1)
                    THROW 51305,N'[HYT_VALIDACION] La firma está inactiva.',1;
                SET @Predeterminada=1;
            END
            ELSE
            BEGIN
                UPDATE dbo.FirmaPresupuestoHYT SET Activo=@Activo,
                    Predeterminada=CASE WHEN @Activo=0 THEN 0 ELSE Predeterminada END WHERE FirmaId=@FirmaId;
                SET @Predeterminada=0;
                UPDATE dbo.ConfiguracionHYT SET FirmaPresupuestoJson=(
                    SELECT firma.ArchivoJson FROM dbo.FirmaPresupuestoHYT AS firma WHERE firma.Predeterminada=1 AND firma.Activo=1
                ),FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=1;
            END;
        END;
        IF @Predeterminada=1
        BEGIN
            UPDATE dbo.FirmaPresupuestoHYT SET Predeterminada=0 WHERE Predeterminada=1 AND FirmaId<>@FirmaId;
            UPDATE dbo.FirmaPresupuestoHYT SET Predeterminada=1 WHERE FirmaId=@FirmaId;
            UPDATE dbo.ConfiguracionHYT SET FirmaPresupuestoJson=@Datos,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=1;
        END;
        DECLARE @Operacion UNIQUEIDENTIFIER=NEWID(),@Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@FirmaId);
        DECLARE @Detalle NVARCHAR(MAX)=(SELECT @Nombre AS Nombre,@Predeterminada AS Predeterminada FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'FirmaPresupuestoHYT',@Registro,N'GUARDAR',@CambiosJson=@Detalle;
        COMMIT TRANSACTION;
        SELECT @FirmaId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_FirmaPresupuestoHYT_Guardar TO hyt_app_executor;
GO
