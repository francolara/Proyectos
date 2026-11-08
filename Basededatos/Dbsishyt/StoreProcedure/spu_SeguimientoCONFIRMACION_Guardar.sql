-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda Confirmación del cliente con idempotencia, concurrencia, archivos y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Claves de adjuntos por carpeta y número, compatibles con cargas históricas. Confirmación vigente completa automáticamente la OC registrada con auditoría transaccional, sin degradar una OC cerrada.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Distingue entre permisos de registro y edición del documento.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PadreId BIGINT,@VersionPadre BINARY(8),@Datos NVARCHAR(MAX),
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'CONFIRMACIONES.CREAR' ELSE N'CONFIRMACIONES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaConfirmacion'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaConfirmacion'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaConfirmacion.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaConfirmacion'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaConfirmacion.',1;
        DECLARE @FechaConfirmacion DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaConfirmacion'),N''));
        IF @FechaConfirmacion IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaConfirmacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'NumeroConfirmacion' AND LEN(campo.value)>100) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: NumeroConfirmacion.',1;
        DECLARE @NumeroConfirmacion NVARCHAR(100)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroConfirmacion'))),N'');
        IF @NumeroConfirmacion IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: NumeroConfirmacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'RepresentanteCliente' AND LEN(campo.value)>150) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: RepresentanteCliente.',1;
        DECLARE @RepresentanteCliente NVARCHAR(150)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RepresentanteCliente'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Observacion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Observacion.',1;
        DECLARE @Observacion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Observacion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'CONFIRMACION_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'CONFIRMACION_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
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
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'VIGENTE',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @PresupuestoId=oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        IF @PresupuestoId IS NULL THROW 51410,N'[HYT_VALIDACION] La OC del cliente no existe.',1;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre AND oc.Estado<>N'CANCELADA')
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió o está cancelada. Vuelve a consultarla.',1;
        SELECT TOP (1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de modificar documentos del expediente.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.FechaConfirmacion, documento.NumeroConfirmacion, documento.RepresentanteCliente, documento.Observacion, documento.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END
            FROM dbo.ConfirmacionFacturacion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.ConfirmacionId=@Id AND documento.OrdenCompraId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.ConfirmacionFacturacion AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.OrdenCompraId=@PadreId AND documento.Vigente=1) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        IF @Estado=N'RETIRADA' THROW 51415,N'[HYT_VALIDACION] La confirmación retirada se conserva como historial. Registra una nueva.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ConfirmacionFacturacion (FechaConfirmacion,NumeroConfirmacion,RepresentanteCliente,Observacion,OrdenCompraId,UsuarioRegistro) VALUES (@FechaConfirmacion,@NumeroConfirmacion,@RepresentanteCliente,@Observacion,@PadreId,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.ConfirmacionFacturacion SET FechaConfirmacion=@FechaConfirmacion,NumeroConfirmacion=@NumeroConfirmacion,RepresentanteCliente=@RepresentanteCliente,Observacion=@Observacion,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ConfirmacionId=@Id;
        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo WHERE vinculo.ConfirmacionId=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.ConfirmacionArchivo AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.ConfirmacionId=@Id
        WHERE archivo.Retirado=0 AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE TRY_CONVERT(BIGINT,entrada.value)=archivo.ArchivoId);
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoAnterior,EstadoNuevo,UsuarioId,Motivo,CambiosJson)
        SELECT @CorrelacionId,N'ArchivoAdjunto',CONVERT(NVARCHAR(100),retirado.ArchivoId),N'RETIRAR',N'VIGENTE',N'RETIRADO',@UsuarioId,@Motivo,
            (SELECT retirado.NombreOriginal AS NombreOriginal FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)
        FROM @ArchivosRetirados AS retirado;
        IF @ArchivoJson IS NOT NULL
        BEGIN
            DECLARE @Nombre NVARCHAR(255)=JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'),@Clave NVARCHAR(500)=JSON_VALUE(@ArchivoJson,N'$.ClaveAlmacenamiento'),@TipoContenido NVARCHAR(150)=JSON_VALUE(@ArchivoJson,N'$.TipoContenido'),
                @Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@ArchivoJson,N'$.TamanoBytes')),@SHA CHAR(64)=JSON_VALUE(@ArchivoJson,N'$.SHA256'),@Extension NVARCHAR(5);
            SET @Extension=CASE @TipoContenido WHEN N'application/pdf' THEN N'.pdf' WHEN N'image/png' THEN N'.png' WHEN N'image/jpeg' THEN N'.jpg' END;
            IF NULLIF(@Nombre,N'') IS NULL OR @Tamano IS NULL OR @Tamano<=0 OR @Tamano>15728640 OR @Extension IS NULL OR LEN(JSON_VALUE(@ArchivoJson,N'$.NombreOriginal'))>255
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR (@Clave COLLATE Latin1_General_100_BIN2<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension AND @Clave COLLATE Latin1_General_100_BIN2<>N'ConfirmacionCliente/'+dbo.fn_ArchivoSegmento(@NumeroConfirmacion)+N'/'+LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension)
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.ConfirmacionArchivo (ConfirmacionId,ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
        SELECT @Nuevo=(SELECT actual.FechaConfirmacion, actual.NumeroConfirmacion, actual.RepresentanteCliente, actual.Observacion, actual.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfirmacionFacturacion AS actual WHERE actual.ConfirmacionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfirmacionFacturacion',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        DECLARE @EstadoOCAnterior NVARCHAR(40);
        SELECT @EstadoOCAnterior=oc.Estado FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        UPDATE dbo.OrdenCompraCliente
        SET Estado=CASE WHEN Estado=N'REGISTRADA' THEN N'TRABAJO_COMPLETADO' ELSE Estado END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        WHERE OrdenCompraId=@PadreId;
        IF @EstadoOCAnterior=N'REGISTRADA'
        BEGIN
            DECLARE @RegistroOC NVARCHAR(100)=CONVERT(NVARCHAR(100),@PadreId),
                @MotivoOC NVARCHAR(2000)=N'Trabajo completado automáticamente por confirmación del cliente.',
                @CambiosOC NVARCHAR(MAX)=(SELECT @Id AS ConfirmacionId,@NumeroConfirmacion AS NumeroConfirmacion,@RepresentanteCliente AS RepresentanteCliente,
                    JSON_QUERY(N'{"Estado":"REGISTRADA"}') AS Anterior,JSON_QUERY(N'{"Estado":"TRABAJO_COMPLETADO"}') AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'OrdenCompraCliente',@RegistroOC,N'CAMBIAR_ESTADO_AUTOMATICO',@EstadoOCAnterior,N'TRABAJO_COMPLETADO',@FechaConfirmacion,@MotivoOC,@CambiosOC;
        END;
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
