-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda Orden de Compra - Cliente con idempotencia, concurrencia, archivos y auditoría.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_Guardar
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
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.RevisionId'),N'') IS NOT NULL AND TRY_CONVERT(BIGINT,NULLIF(JSON_VALUE(@Datos,N'$.RevisionId'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: RevisionId.',1;
        DECLARE @RevisionId BIGINT=TRY_CONVERT(BIGINT,NULLIF(JSON_VALUE(@Datos,N'$.RevisionId'),N''));
        IF @RevisionId IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: RevisionId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'NumeroOCCliente' AND LEN(campo.value)>100) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: NumeroOCCliente.',1;
        DECLARE @NumeroOCCliente NVARCHAR(100)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroOCCliente'))),N'');
        IF @NumeroOCCliente IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: NumeroOCCliente.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaEmisionOC'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaEmisionOC'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaEmisionOC.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaEmisionOC'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaEmisionOC.',1;
        DECLARE @FechaEmisionOC DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaEmisionOC'),N''));
        IF @FechaEmisionOC IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaEmisionOC.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaRecepcionOC.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaRecepcionOC.',1;
        DECLARE @FechaRecepcionOC DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaRecepcionOC'),N''));
        IF @FechaRecepcionOC IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaRecepcionOC.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.Importe'),N'') IS NOT NULL AND TRY_CONVERT(DECIMAL(19,2),NULLIF(JSON_VALUE(@Datos,N'$.Importe'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: Importe.',1;
        DECLARE @Importe DECIMAL(19,2)=TRY_CONVERT(DECIMAL(19,2),NULLIF(JSON_VALUE(@Datos,N'$.Importe'),N''));
        IF @Importe<0 OR (JSON_VALUE(@Datos,N'$.Importe') IS NOT NULL AND (PATINDEX(N'%[^0-9.]%',JSON_VALUE(@Datos,N'$.Importe'))>0 OR (CHARINDEX(N'.',JSON_VALUE(@Datos,N'$.Importe'))>0 AND LEN(JSON_VALUE(@Datos,N'$.Importe'))-CHARINDEX(N'.',JSON_VALUE(@Datos,N'$.Importe'))>2))) THROW 51408,N'[HYT_VALIDACION] Importe no negativo con hasta dos decimales.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Observacion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Observacion.',1;
        DECLARE @Observacion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Observacion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'OC_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'OC_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
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
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'REGISTRADA',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
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
            SELECT @Anterior=(SELECT documento.RevisionId, documento.NumeroOCCliente, documento.FechaEmisionOC, documento.FechaRecepcionOC, documento.Importe, documento.Observacion, documento.NumeroRegistroInterno, documento.CodigoMoneda, documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=documento.Estado
            FROM dbo.OrdenCompraCliente AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.OrdenCompraId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.PresupuestoId=@PadreId) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        DECLARE @EstadoRevision NVARCHAR(20),@TotalRevision DECIMAL(19,2);
        SELECT @EstadoRevision=r.Estado,@TotalRevision=r.Total,@Moneda=r.CodigoMoneda FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId AND r.RevisionId=@RevisionId;
        IF @EstadoRevision IS NULL OR @TotalRevision IS NULL OR (@TipoAtencion=N'INVITACION' AND @EstadoRevision<>N'APROBADO') OR (@TipoAtencion=N'EMERGENCIA' AND @EstadoRevision NOT IN (N'BORRADOR',N'ENVIADO',N'APROBADO'))
            THROW 51414,N'[HYT_VALIDACION] Invitación requiere revisión aprobada. Emergencia admite revisión completa sin aprobación previa.',1;
        IF @Estado=N'CANCELADA' THROW 51415,N'[HYT_VALIDACION] Reactiva la OC antes de editarla.',1;
        IF @Id IS NOT NULL AND (@Estado=N'CERRADA_OPERATIVAMENTE' OR EXISTS (SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1))
            AND (@RevisionId<>TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.RevisionId')) OR @NumeroOCCliente<>JSON_VALUE(@Anterior,N'$.NumeroOCCliente')
                OR @FechaEmisionOC<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaEmisionOC'),23) OR @FechaRecepcionOC<>TRY_CONVERT(DATE,JSON_VALUE(@Anterior,N'$.FechaRecepcionOC'),23)
                OR COALESCE(@Importe,-1)<>COALESCE(TRY_CONVERT(DECIMAL(19,2),JSON_VALUE(@Anterior,N'$.Importe')),-1))
            THROW 51415,N'[HYT_VALIDACION] Reabre la OC o retira su confirmación antes de corregir sus datos principales.',1;
        DECLARE @NumeroInterno NVARCHAR(40);
        IF @Id IS NULL EXEC dbo.spu_ObtenerNumeroInterno @TipoRegistro=N'OC_CLIENTE',@NumeroInterno=@NumeroInterno OUTPUT;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.OrdenCompraCliente (RevisionId,NumeroOCCliente,FechaEmisionOC,FechaRecepcionOC,Importe,Observacion,PresupuestoId,UsuarioRegistro,NumeroRegistroInterno,CodigoMoneda) VALUES (@RevisionId,@NumeroOCCliente,@FechaEmisionOC,@FechaRecepcionOC,@Importe,@Observacion,@PadreId,@UsuarioId,@NumeroInterno,@Moneda);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.OrdenCompraCliente SET RevisionId=@RevisionId,NumeroOCCliente=@NumeroOCCliente,FechaEmisionOC=@FechaEmisionOC,FechaRecepcionOC=@FechaRecepcionOC,Importe=@Importe,Observacion=@Observacion,CodigoMoneda=@Moneda,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@Id;
        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo WHERE vinculo.OrdenCompraId=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.OrdenCompraArchivo AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.OrdenCompraId=@Id
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
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR @Clave<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.OrdenCompraArchivo (OrdenCompraId,ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
        SELECT @Nuevo=(SELECT actual.RevisionId, actual.NumeroOCCliente, actual.FechaEmisionOC, actual.FechaRecepcionOC, actual.Importe, actual.Observacion, actual.NumeroRegistroInterno, actual.CodigoMoneda, actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.OrdenCompraCliente AS actual WHERE actual.OrdenCompraId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'OrdenCompraCliente',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
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
