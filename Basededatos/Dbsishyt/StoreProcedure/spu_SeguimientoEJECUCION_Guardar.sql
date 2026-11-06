-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda Ejecución del trabajo con idempotencia, concurrencia, archivos y auditoría.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoEJECUCION_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'EJECUCION.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Responsable' AND LEN(campo.value)>150) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Responsable.',1;
        DECLARE @Responsable NVARCHAR(150)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Responsable'))),N'');
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaProgramada'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaProgramada'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaProgramada.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaProgramada'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaProgramada.',1;
        DECLARE @FechaProgramada DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaProgramada'),N''));
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaInicioReal'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaInicioReal'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaInicioReal.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaInicioReal'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaInicioReal.',1;
        DECLARE @FechaInicioReal DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaInicioReal'),N''));
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaTerminoReal'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaTerminoReal'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaTerminoReal.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaTerminoReal'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaTerminoReal.',1;
        DECLARE @FechaTerminoReal DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaTerminoReal'),N''));
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Descripcion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Descripcion.',1;
        DECLARE @Descripcion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'EJECUCION_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'EJECUCION_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Resultado BIGINT,@ArchivoId BIGINT,@ActorAnterior NVARCHAR(128),@HuellaAnterior CHAR(64),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@ArchivoId=operacion.ArchivoId,@ActorAnterior=operacion.UsuarioId,@HuellaAnterior=operacion.Huella,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH (UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @ActorAnterior<>@UsuarioId OR @HuellaAnterior<>@Huella OR @TipoAnterior<>@TipoOperacion OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] Este identificador de operación ya se usó con otros datos. Vuelve a abrir el formulario.',1;
            COMMIT TRANSACTION;
            SELECT @Resultado AS Id,@ArchivoId AS ArchivoId;
            RETURN;
        END;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'PENDIENTE',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SET @PresupuestoId=@PadreId;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP (1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de modificar documentos del expediente.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.Responsable, documento.FechaProgramada, documento.FechaInicioReal, documento.FechaTerminoReal, documento.Descripcion, documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=documento.Estado
            FROM dbo.Ejecucion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.EjecucionId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.Ejecucion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.PresupuestoId=@PadreId) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        IF @FechaTerminoReal IS NOT NULL AND (@FechaInicioReal IS NULL OR @FechaTerminoReal<@FechaInicioReal) THROW 51417,N'[HYT_VALIDACION] Revisa inicio y término reales.',1;
        IF @Estado=N'CANCELADA' THROW 51415,N'[HYT_VALIDACION] Reactiva la ejecución antes de editarla.',1;
        IF @Estado=N'PROGRAMADA' AND @FechaProgramada IS NULL THROW 51417,N'[HYT_VALIDACION] La ejecución programada necesita fecha.',1;
        IF @Estado IN (N'EN_EJECUCION',N'TERMINADA') AND (@FechaInicioReal IS NULL OR @Responsable IS NULL OR (@Estado=N'TERMINADA' AND @FechaTerminoReal IS NULL))
            THROW 51417,N'[HYT_VALIDACION] Conserva responsable y fechas reales según el estado.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Ejecucion (Responsable,FechaProgramada,FechaInicioReal,FechaTerminoReal,Descripcion,PresupuestoId,UsuarioRegistro) VALUES (@Responsable,@FechaProgramada,@FechaInicioReal,@FechaTerminoReal,@Descripcion,@PadreId,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.Ejecucion SET Responsable=@Responsable,FechaProgramada=@FechaProgramada,FechaInicioReal=@FechaInicioReal,FechaTerminoReal=@FechaTerminoReal,Descripcion=@Descripcion,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE EjecucionId=@Id;
        IF JSON_QUERY(@Datos,N'$.Adjunto') IS NOT NULL THROW 51419,N'[HYT_VALIDACION] Esta ficha no admite adjuntos.',1;
        SELECT @Nuevo=(SELECT actual.Responsable, actual.FechaProgramada, actual.FechaInicioReal, actual.FechaTerminoReal, actual.Descripcion, actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.Ejecucion AS actual WHERE actual.EjecucionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Ejecucion',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion (OperacionId,TipoOperacion,PadreId,RegistroId,ArchivoId,Huella,UsuarioId)
        VALUES (@CorrelacionId,@TipoOperacion,@PadreId,@Id,@ArchivoId,@Huella,@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Id AS Id,@ArchivoId AS ArchivoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
