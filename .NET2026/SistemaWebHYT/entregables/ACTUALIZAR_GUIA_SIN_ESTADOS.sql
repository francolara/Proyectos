-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Guía como documento registrado sin estados; conserva historial, datos, adjuntos y requisitos de ejecución/confirmación.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Facturado admite OC completada con guía y confirmación vigente. Reportes integrales y dashboard alineados al flujo vigente. Confirmación del cliente y Facturado por registros vigentes.
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
GO

-- Fuente: Script/007_ProteccionAuditoria.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GO
CREATE OR ALTER TRIGGER dbo.tr_EventoAuditoria_Inmutable
ON dbo.EventoAuditoria
INSTEAD OF UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;
    THROW 51009, N'El historial de auditoria es inmutable.', 1;
END;
GO

CREATE OR ALTER TRIGGER dbo.tr_TasaIGV_SinSolapamiento
ON dbo.TasaIGV
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM dbo.TasaIGV AS primera WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN dbo.TasaIGV AS segunda WITH (UPDLOCK, HOLDLOCK)
            ON primera.FechaDesde < segunda.FechaDesde
            AND (primera.FechaHasta IS NULL OR segunda.FechaDesde < primera.FechaHasta)
    )
        THROW 51010, N'Las vigencias de IGV no pueden solaparse.', 1;
END;
GO

CREATE OR ALTER TRIGGER dbo.tr_ReferenciaComprobante_Requisitos
ON dbo.ReferenciaComprobanteExterno
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM inserted AS referencia
        INNER JOIN dbo.OrdenCompraCliente AS oc WITH (UPDLOCK, HOLDLOCK) ON oc.OrdenCompraId = referencia.OrdenCompraId
        INNER JOIN dbo.ConfirmacionFacturacion AS confirmacion WITH (UPDLOCK, HOLDLOCK)
            ON confirmacion.ConfirmacionId = referencia.ConfirmacionId
        WHERE referencia.Vigente = 1 AND
        (
            confirmacion.Vigente = 0 OR oc.Estado = N'CANCELADA'
            OR NOT EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS guia WITH (UPDLOCK, HOLDLOCK)
                WHERE guia.PresupuestoId = oc.PresupuestoId)
            OR confirmacion.OrdenCompraId <> oc.OrdenCompraId
            OR (oc.Estado NOT IN (N'TRABAJO_COMPLETADO', N'CERRADA_OPERATIVAMENTE') AND
                (NOT EXISTS (SELECT 1 FROM dbo.Ejecucion AS ejecucion WITH (UPDLOCK, HOLDLOCK)
                    WHERE ejecucion.PresupuestoId = oc.PresupuestoId AND ejecucion.Estado = N'TERMINADA')
                OR EXISTS (SELECT 1 FROM dbo.Ejecucion AS pendiente WITH (UPDLOCK, HOLDLOCK)
                    WHERE pendiente.PresupuestoId = oc.PresupuestoId AND pendiente.Estado NOT IN (N'TERMINADA', N'CANCELADA'))))
        )
    )
        THROW 51013, N'El marcado Facturado requiere OC valida, trabajo terminado, guia registrada y confirmacion vigente.', 1;
END;
GO
CREATE OR ALTER TRIGGER dbo.tr_Confirmacion_ProtegerFacturado
ON dbo.ConfirmacionFacturacion
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM inserted AS confirmacion
        INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia WITH (UPDLOCK, HOLDLOCK)
            ON referencia.ConfirmacionId = confirmacion.ConfirmacionId AND referencia.Vigente = 1
        WHERE confirmacion.Vigente = 0
    )
        THROW 51014, N'Retire primero el marcado Facturado con motivo y auditoria antes de revocar su confirmacion.', 1;
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoGUIA_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_Guardar
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
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'GUIAS.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @PadreId IS NULL OR @PadreId<=0 OR @Datos IS NULL OR ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{'
            THROW 51403,N'[HYT_VALIDACION] Datos o identificador de operación no válidos.',1;
        DECLARE @Motivo NVARCHAR(1000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.MotivoCambio'))),N'');
        IF @Id IS NOT NULL AND @Motivo IS NULL THROW 51404,N'[HYT_VALIDACION] Indica el motivo de la modificación.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.MotivoCambio'))>1000 THROW 51404,N'[HYT_VALIDACION] El motivo admite hasta 1000 caracteres.',1;
        IF NULLIF(JSON_VALUE(@Datos,N'$.FechaActa'),N'') IS NOT NULL AND TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaActa'),N'')) IS NULL THROW 51406,N'[HYT_VALIDACION] Formato no válido: FechaActa.',1;
        IF LEN(JSON_VALUE(@Datos,N'$.FechaActa'))>10 THROW 51406,N'[HYT_VALIDACION] Usa una fecha sin hora: FechaActa.',1;
        DECLARE @FechaActa DATE=TRY_CONVERT(DATE,NULLIF(JSON_VALUE(@Datos,N'$.FechaActa'),N''));
        IF @FechaActa IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: FechaActa.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'NumeroActa' AND LEN(campo.value)>100) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: NumeroActa.',1;
        DECLARE @NumeroActa NVARCHAR(100)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroActa'))),N'');
        IF @NumeroActa IS NULL THROW 51407,N'[HYT_VALIDACION] Dato obligatorio: NumeroActa.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'Observacion' AND LEN(campo.value)>2000) THROW 51405,N'[HYT_VALIDACION] Texto demasiado largo: Observacion.',1;
        DECLARE @Observacion NVARCHAR(2000)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Observacion'))),N'');
        DECLARE @TipoOperacion NVARCHAR(40)=N'GUIA_GUARDAR',@Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),CONCAT(CONVERT(NVARCHAR(MAX),N'GUIA_GUARDAR|'),@PadreId,N'|',COALESCE(@Id,0),N'|',CONVERT(NVARCHAR(16),@VersionOriginal,2),N'|',CONVERT(NVARCHAR(16),@VersionPadre,2),N'|',@Datos))),2);
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
        DECLARE @RevisionAceptadaId BIGINT;
        DECLARE @PresupuestoId BIGINT,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20),@Moneda NVARCHAR(3),@Estado NVARCHAR(40)=N'REGISTRADA',@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SET @PresupuestoId=@PadreId;
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP (1) @RevisionAceptadaId=r.RevisionId,@EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH (HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto IS NULL OR @EstadoPresupuesto=N'CANCELADO' THROW 51412,N'[HYT_VALIDACION] El presupuesto no admite este registro.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de modificar documentos del expediente.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT documento.FechaActa, documento.NumeroActa, documento.Observacion, documento.FechaAceptacion, documento.RepresentanteCliente, documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@Estado=documento.Estado
            FROM dbo.GuiaConformidad AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.GuiaId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El registro cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NULL AND EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS documento WITH (UPDLOCK,HOLDLOCK) WHERE documento.PresupuestoId=@PadreId) THROW 51413,N'[HYT_VALIDACION] Ya existe este registro. Ábrelo para consultar o modificar.',1;
        IF @Id IS NULL AND @EstadoPresupuesto NOT IN (N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO')
            THROW 51416,N'[HYT_VALIDACION] El presupuesto no está vigente para registrar la guía.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.GuiaConformidad (FechaActa,NumeroActa,Observacion,PresupuestoId,UsuarioRegistro) VALUES (@FechaActa,@NumeroActa,@Observacion,@PadreId,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.GuiaConformidad SET FechaActa=@FechaActa,NumeroActa=@NumeroActa,Observacion=@Observacion,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE GuiaId=@Id;
        DECLARE @ArchivoJson NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Adjunto'),@Retirados NVARCHAR(MAX)=COALESCE(JSON_QUERY(@Datos,N'$.RetirarAdjuntos'),N'[]');
        IF LEFT(LTRIM(@Retirados),1)<>N'[' THROW 51418,N'[HYT_VALIDACION] Selección de adjuntos no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Retirados) AS entrada WHERE entrada.type<>2 OR TRY_CONVERT(BIGINT,entrada.value) IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.GuiaArchivo AS vinculo WHERE vinculo.GuiaId=@Id AND vinculo.ArchivoId=TRY_CONVERT(BIGINT,entrada.value)))
            THROW 51418,N'[HYT_VALIDACION] Un adjunto no pertenece a este documento.',1;
        IF @Motivo IS NULL AND EXISTS (SELECT 1 FROM OPENJSON(@Retirados)) THROW 51418,N'[HYT_VALIDACION] Indica el motivo del retiro de adjuntos.',1;
        DECLARE @ArchivosRetirados TABLE (ArchivoId BIGINT NOT NULL,NombreOriginal NVARCHAR(255) NOT NULL);
        UPDATE archivo SET Retirado=1,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
        OUTPUT inserted.ArchivoId,inserted.NombreOriginal INTO @ArchivosRetirados (ArchivoId,NombreOriginal)
        FROM dbo.ArchivoAdjunto AS archivo INNER JOIN dbo.GuiaArchivo AS vinculo ON vinculo.ArchivoId=archivo.ArchivoId AND vinculo.GuiaId=@Id
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
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR (@Clave COLLATE Latin1_General_100_BIN2<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension AND @Clave COLLATE Latin1_General_100_BIN2<>N'Guia/'+dbo.fn_ArchivoSegmento(@NumeroActa)+N'/'+LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension)
                THROW 51419,N'[HYT_VALIDACION] Metadatos del adjunto no válidos.',1;
            IF @TipoContenido<>N'application/pdf' THROW 51419,N'[HYT_VALIDACION] La guía solo admite escaneado PDF.',1;
            INSERT INTO dbo.ArchivoAdjunto (NombreOriginal,ClaveAlmacenamiento,TipoContenido,TamanoBytes,SHA256,UsuarioRegistro)
            VALUES (@Nombre,@Clave,@TipoContenido,@Tamano,@SHA,@UsuarioId);
            SET @ArchivoId=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.GuiaArchivo (GuiaId,ArchivoId) VALUES (@Id,@ArchivoId);
            DECLARE @RegistroArchivo NVARCHAR(100)=CONVERT(NVARCHAR(100),@ArchivoId),@DetalleArchivo NVARCHAR(MAX)=(SELECT @Nombre AS NombreOriginal,@TipoContenido AS TipoContenido,@Tamano AS TamanoBytes,@SHA AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ArchivoAdjunto',@RegistroArchivo,N'CREAR',NULL,N'VIGENTE',NULL,NULL,@DetalleArchivo;
        END;
        SELECT @Nuevo=(SELECT actual.FechaActa, actual.NumeroActa, actual.Observacion, actual.FechaAceptacion, actual.RepresentanteCliente, actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.GuiaConformidad AS actual WHERE actual.GuiaId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@ArchivoId AS ArchivoAgregado,JSON_QUERY(JSON_QUERY(@Datos,N'$.RetirarAdjuntos')) AS ArchivosRetirados FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),
            @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoAnterior NVARCHAR(40)=CASE WHEN @Anterior IS NULL THEN NULL ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'GuiaConformidad',@Registro,@Accion,@EstadoAnterior,@Estado,NULL,@Motivo,@Cambios;
        IF @Anterior IS NULL AND @EstadoPresupuesto<>N'APROBADO'
        BEGIN
            DECLARE @FechaAceptacionAutomatica DATE=CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()));
            UPDATE dbo.PresupuestoRevision SET Estado=N'APROBADO',FechaAprobacion=@FechaAceptacionAutomatica,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionAceptadaId;
            DECLARE @RegistroRevision NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionAceptadaId),
                @MotivoAceptacion NVARCHAR(1000)=N'Aceptación automática por registro de GUIA.',
                @DetalleAceptacion NVARCHAR(MAX)=(SELECT @EstadoPresupuesto AS EstadoAnterior,N'APROBADO' AS EstadoNuevo,
                    @FechaAceptacionAutomatica AS FechaEvento,N'GUIA' AS TipoDocumento,@Id AS DocumentoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@RegistroRevision,N'CAMBIAR_ESTADO',
                @EstadoPresupuesto,N'APROBADO',NULL,@MotivoAceptacion,@DetalleAceptacion;
        END;
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


-- Fuente: StoreProcedure/spu_SeguimientoGUIA_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
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


-- Fuente: StoreProcedure/spu_SeguimientoOC_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'OC_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=@PadreId,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=documento.Estado,@Anterior=(SELECT documento.RevisionId,documento.NumeroOCCliente,documento.FechaEmisionOC,documento.FechaRecepcionOC,documento.Importe,documento.Observacion,documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.OrdenCompraCliente AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.OrdenCompraId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'REGISTRADA' AND @NuevoEstado=N'TRABAJO_COMPLETADO') OR (@EstadoAnterior=N'REGISTRADA' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'CERRADA_OPERATIVAMENTE' AND @NuevoEstado=N'TRABAJO_COMPLETADO') OR (@EstadoAnterior=N'CANCELADA' AND @NuevoEstado=N'REGISTRADA')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        IF @NuevoEstado=N'TRABAJO_COMPLETADO'
            AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA')
            AND NOT EXISTS(SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1)
            THROW 51433,N'[HYT_VALIDACION] Registra la confirmación del cliente o la ejecución terminada para completar el trabajo.',1;
        IF @NuevoEstado=N'CERRADA_OPERATIVAMENTE' AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA')
            THROW 51433,N'[HYT_VALIDACION] Para cerrar la OC registra la ejecución terminada del trabajo.',1;
        IF @NuevoEstado=N'CERRADA_OPERATIVAMENTE' AND NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId)
            THROW 51433,N'[HYT_VALIDACION] Para cerrar la OC se requiere la guía registrada.',1;
        IF @NuevoEstado=N'CANCELADA' AND (EXISTS(SELECT 1 FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.OrdenCompraId=@Id AND confirmacion.Vigente=1)
            OR EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'EN_EJECUCION'))
            THROW 51433,N'[HYT_VALIDACION] Retira la confirmación, detén la ejecución antes de cancelar la OC.',1;
        UPDATE dbo.OrdenCompraCliente SET Estado=@NuevoEstado,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@Id;
        SELECT @Nuevo=(SELECT actual.RevisionId,actual.NumeroOCCliente,actual.FechaEmisionOC,actual.FechaRecepcionOC,actual.Importe,actual.Observacion,actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.OrdenCompraCliente AS actual WHERE actual.OrdenCompraId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'OrdenCompraCliente',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_SeguimientoEJECUCION_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoEJECUCION_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'EJECUCION.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'EJECUCION_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=@PadreId,@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PadreId AND p.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=documento.Estado,@Anterior=(SELECT documento.Responsable,documento.FechaProgramada,documento.FechaInicioReal,documento.FechaTerminoReal,documento.Descripcion,documento.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Ejecucion AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.EjecucionId=@Id AND documento.PresupuestoId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'PENDIENTE' AND @NuevoEstado=N'PROGRAMADA') OR (@EstadoAnterior=N'PENDIENTE' AND @NuevoEstado=N'EN_EJECUCION') OR (@EstadoAnterior=N'PENDIENTE' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'PROGRAMADA' AND @NuevoEstado=N'PENDIENTE') OR (@EstadoAnterior=N'PROGRAMADA' AND @NuevoEstado=N'EN_EJECUCION') OR (@EstadoAnterior=N'PROGRAMADA' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'EN_EJECUCION' AND @NuevoEstado=N'TERMINADA') OR (@EstadoAnterior=N'EN_EJECUCION' AND @NuevoEstado=N'CANCELADA') OR (@EstadoAnterior=N'TERMINADA' AND @NuevoEstado=N'EN_EJECUCION') OR (@EstadoAnterior=N'CANCELADA' AND @NuevoEstado=N'PENDIENTE')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        IF @NuevoEstado=N'PROGRAMADA' AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND ejecucion.FechaProgramada IS NOT NULL AND NULLIF(LTRIM(RTRIM(ejecucion.Responsable)),N'') IS NOT NULL)
            THROW 51433,N'[HYT_VALIDACION] Completa fecha programada y responsable del trabajo en la ficha.',1;
        IF @NuevoEstado=N'EN_EJECUCION' AND (NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND NULLIF(LTRIM(RTRIM(ejecucion.Responsable)),N'') IS NOT NULL)
            OR (@TipoAtencion=N'INVITACION' AND (NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado<>N'CANCELADA')
                OR NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND ejecucion.FechaProgramada IS NOT NULL))))
            THROW 51433,N'[HYT_VALIDACION] Completa responsable; invitación requiere OC y programación antes de iniciar.',1;
        IF @NuevoEstado=N'TERMINADA' AND EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.EjecucionId=@Id AND (ejecucion.FechaInicioReal IS NULL OR CONVERT(DATE,ejecucion.FechaInicioReal)>@FechaEvento))
            THROW 51433,N'[HYT_VALIDACION] La fecha de término debe ser igual o posterior al inicio real.',1;
        IF @EstadoAnterior=N'TERMINADA' AND (EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@PresupuestoId AND oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE')))
            THROW 51433,N'[HYT_VALIDACION] Reabre la OC antes de reabrir el trabajo.',1;
        UPDATE dbo.Ejecucion SET Estado=@NuevoEstado,
            FechaInicioReal=CASE WHEN @NuevoEstado=N'EN_EJECUCION' THEN COALESCE(FechaInicioReal,CONVERT(DATETIME2(3),@FechaEvento)) ELSE FechaInicioReal END,
            FechaTerminoReal=CASE WHEN @NuevoEstado=N'TERMINADA' THEN CONVERT(DATETIME2(3),@FechaEvento) WHEN @NuevoEstado=N'EN_EJECUCION' THEN NULL ELSE FechaTerminoReal END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE EjecucionId=@Id;
        SELECT @Nuevo=(SELECT actual.Responsable,actual.FechaProgramada,actual.FechaInicioReal,actual.FechaTerminoReal,actual.Descripcion,actual.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.Ejecucion AS actual WHERE actual.EjecucionId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Ejecucion',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Presupuesto_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PresupuestoId BIGINT,@RevisionId BIGINT,
    @VersionOriginal BINARY(8),@VersionPresupuesto BINARY(8),@NuevoEstado NVARCHAR(20),
    @FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para cambiar el estado.',1;
        IF @FechaEvento IS NULL OR NULLIF(LTRIM(RTRIM(@Responsable)),N'') IS NULL OR @NuevoEstado IS NULL
            THROW 51318,N'[HYT_VALIDACION] Indica fecha y responsable de la acción.',1;
        IF @NuevoEstado IN (N'OBSERVADO',N'RECHAZADO',N'CANCELADO') AND NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL
            THROW 51318,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId AND p.VersionFila=@VersionPresupuesto)
            THROW 51309,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        DECLARE @Estado NVARCHAR(20),@Empresa NVARCHAR(MAX),@Moneda NVARCHAR(3),@Total DECIMAL(19,2);
        SELECT @Estado=r.Estado,@Empresa=r.EmpresaSnapshotJson,@Moneda=r.CodigoMoneda,@Total=r.Total
        FROM dbo.PresupuestoRevision AS r WITH (UPDLOCK,HOLDLOCK)
        WHERE r.RevisionId=@RevisionId AND r.PresupuestoId=@PresupuestoId AND r.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió o no pertenece al presupuesto.',1;
        IF EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.NumeroRevision>(SELECT origen.NumeroRevision FROM dbo.PresupuestoRevision AS origen WHERE origen.RevisionId=@RevisionId))
            THROW 51319,N'[HYT_VALIDACION] Solo se cambia el estado de la última revisión.',1;
        IF NOT ((@Estado IN (N'BORRADOR',N'OBSERVADO') AND @NuevoEstado IN (N'ENVIADO',N'CANCELADO'))
            OR (@Estado=N'ENVIADO' AND @NuevoEstado IN (N'OBSERVADO',N'APROBADO',N'RECHAZADO',N'CANCELADO'))
            OR (@Estado IN (N'RECHAZADO',N'VENCIDO',N'APROBADO') AND @NuevoEstado=N'CANCELADO'))
            THROW 51319,N'[HYT_VALIDACION] La transición no está permitida. Para corregir contenido crea otra revisión.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (HOLDLOCK) WHERE oc.PresupuestoId=@PresupuestoId)
            AND NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId
                WHERE oc.PresupuestoId=@PresupuestoId AND oc.RevisionId=@RevisionId AND oc.Estado<>N'CANCELADA' AND p.TipoAtencion=N'EMERGENCIA'
                AND ((@Estado=N'BORRADOR' AND @NuevoEstado=N'ENVIADO') OR (@Estado=N'ENVIADO' AND @NuevoEstado=N'APROBADO')))
            THROW 51320,N'[HYT_VALIDACION] La revisión vinculada a una OC no permite este cambio.',1;
        IF @NuevoEstado=N'ENVIADO'
        BEGIN
            IF @Total IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS linea WHERE linea.RevisionId=@RevisionId)
                OR EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS linea WHERE linea.RevisionId=@RevisionId AND (linea.Total IS NULL OR linea.PrecioUnitario IS NULL OR linea.CodigoMonedaOrigen<>@Moneda))
                THROW 51321,N'[HYT_VALIDACION] Completa las partidas y los precios en la moneda elegida antes de enviar.',1;
            IF ISJSON(@Empresa)<>1 OR @Empresa IS NULL OR NULLIF(JSON_VALUE(@Empresa,N'$.Ruc'),N'') IS NULL OR NULLIF(JSON_VALUE(@Empresa,N'$.RazonSocial'),N'') IS NULL
                THROW 51321,N'[HYT_VALIDACION] Configura los datos de HYT y vuelve a guardar el borrador antes de enviar.',1;
        END;
        UPDATE dbo.PresupuestoRevision SET Estado=@NuevoEstado,FechaAprobacion=CASE WHEN @NuevoEstado=N'APROBADO' THEN @FechaEvento ELSE FechaAprobacion END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionId;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),@FechaReal DATETIME2(3)=CONVERT(DATETIME2(3),@FechaEvento),
            @MotivoAudit NVARCHAR(1000)=LEFT(@Motivo,1000),@Detalle NVARCHAR(MAX);
        SET @Detalle=(SELECT @Estado AS EstadoAnterior,@NuevoEstado AS EstadoNuevo,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Observacion FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,N'CAMBIAR_ESTADO',@Estado,@NuevoEstado,@FechaReal,@MotivoAudit,@Detalle;
        COMMIT TRANSACTION;
        SELECT p.PresupuestoId,r.RevisionId,p.VersionFila AS VersionPresupuesto,r.VersionFila AS VersionRevision,r.Estado
        FROM dbo.Presupuesto AS p INNER JOIN dbo.PresupuestoRevision AS r ON r.PresupuestoId=p.PresupuestoId AND r.RevisionId=@RevisionId WHERE p.PresupuestoId=@PresupuestoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Facturado_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Facturado_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@OrdenCompraId BIGINT,@VersionOC BINARY(8),
    @FechaDocumento DATE=NULL,@CodigoTipoDocumento NVARCHAR(2)=NULL,@Serie NVARCHAR(30)=NULL,@NumeroDocumento NVARCHAR(50)=NULL,
    @ReferenciaId BIGINT=NULL,@VersionReferencia BINARY(8)=NULL,@Retirar BIT=0,@Motivo NVARCHAR(1000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'FACTURADO.EDITAR')
            THROW 51500,N'[HYT_PERMISO] No tiene permiso para esta referencia de facturación.',1;
        IF @@TRANCOUNT<>0 THROW 51501,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Serie=NULLIF(LTRIM(RTRIM(@Serie)),N'');SET @NumeroDocumento=NULLIF(LTRIM(RTRIM(@NumeroDocumento)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @OrdenCompraId IS NULL OR @OrdenCompraId<=0 OR @VersionOC IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000' OR @Retirar IS NULL
            THROW 51502,N'[HYT_VALIDACION] Vuelve a abrir la OC.',1;
        IF @Retirar=0 AND (@FechaDocumento IS NULL OR @CodigoTipoDocumento IS NULL OR @CodigoTipoDocumento NOT IN(N'01',N'03') OR @Serie IS NULL OR @NumeroDocumento IS NULL)
            THROW 51502,N'[HYT_VALIDACION] Completa fecha, tipo 01/03, serie y número del documento externo.',1;
        IF (@ReferenciaId IS NOT NULL OR @Retirar=1) AND @Motivo IS NULL THROW 51502,N'[HYT_VALIDACION] Indica el motivo de la corrección o retiro.',1;
        IF @Retirar=1 AND @ReferenciaId IS NULL THROW 51502,N'[HYT_VALIDACION] Selecciona la referencia a retirar.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @OrdenCompraId AS OrdenCompraId,@VersionOC AS VersionOC,@ReferenciaId AS ReferenciaId,@VersionReferencia AS VersionReferencia,@FechaDocumento AS FechaDocumento,@CodigoTipoDocumento AS Tipo,@Serie AS Serie,@NumeroDocumento AS Numero,@Retirar AS Retirar,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Resultado BIGINT,@Actor NVARCHAR(128),@HashPrevio CHAR(64),@TipoPrevio NVARCHAR(40),@PadrePrevio BIGINT;
        BEGIN TRANSACTION;SET @Propia=1;
        SELECT @Resultado=operacion.RegistroId,@Actor=operacion.UsuarioId,@HashPrevio=operacion.Huella,@TipoPrevio=operacion.TipoOperacion,@PadrePrevio=operacion.PadreId FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Actor<>@UsuarioId OR @HashPrevio<>@Huella OR @TipoPrevio<>N'FACTURADO_GUARDAR' OR @PadrePrevio<>@OrdenCompraId THROW 51503,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS ReferenciaId;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=(SELECT oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@OrdenCompraId),@Bloqueo BIGINT;
        SELECT @Bloqueo=p.PresupuestoId FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @Bloqueo IS NULL THROW 51504,N'[HYT_VALIDACION] La OC no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH(UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@OrdenCompraId AND oc.VersionFila=@VersionOC AND oc.Estado<>N'CANCELADA')
            THROW 51505,N'[HYT_CONCURRENCIA] La OC cambió o está cancelada. Vuelve a consultarla.',1;
        DECLARE @ConfirmacionId BIGINT,@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@EstadoAnterior NVARCHAR(30);
        SELECT @ConfirmacionId=confirmacion.ConfirmacionId FROM dbo.ConfirmacionFacturacion AS confirmacion WITH(UPDLOCK,HOLDLOCK) WHERE confirmacion.OrdenCompraId=@OrdenCompraId AND confirmacion.Vigente=1;
        IF @Retirar=0 AND @ConfirmacionId IS NULL THROW 51506,N'[HYT_VALIDACION] Registra primero la confirmación del cliente de esta OC.',1;
        IF @Retirar=0 AND (NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@PresupuestoId) OR (NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@OrdenCompraId AND oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE')) AND NOT EXISTS(SELECT 1 FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@PresupuestoId AND ejecucion.Estado=N'TERMINADA')))
            THROW 51506,N'[HYT_VALIDACION] Para marcar Facturado la OC debe tener trabajo completado (o ejecución terminada) y una guía registrada.',1;
        IF @ReferenciaId IS NULL AND EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS referencia WITH(UPDLOCK,HOLDLOCK) WHERE referencia.OrdenCompraId=@OrdenCompraId AND referencia.Vigente=1)
            THROW 51507,N'[HYT_VALIDACION] La OC ya está marcada. Abre su referencia para corregirla.',1;
        IF @ReferenciaId IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.Vigente,referencia.ConfirmacionId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.ReferenciaComprobanteExterno AS referencia WITH(UPDLOCK,HOLDLOCK) WHERE referencia.ReferenciaId=@ReferenciaId AND referencia.OrdenCompraId=@OrdenCompraId AND referencia.VersionFila=@VersionReferencia AND referencia.Vigente=1;
            IF @Anterior IS NULL THROW 51505,N'[HYT_CONCURRENCIA] La referencia cambió. Vuelve a consultarla.',1;
            SET @EstadoAnterior=N'FACTURADO';
        END;
        IF @ReferenciaId IS NULL
        BEGIN
            INSERT INTO dbo.ReferenciaComprobanteExterno(OrdenCompraId,ConfirmacionId,FechaDocumento,CodigoTipoDocumento,Serie,NumeroDocumento,UsuarioRegistro) VALUES(@OrdenCompraId,@ConfirmacionId,@FechaDocumento,@CodigoTipoDocumento,@Serie,@NumeroDocumento,@UsuarioId);
            SET @ReferenciaId=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE IF @Retirar=1
            UPDATE dbo.ReferenciaComprobanteExterno SET Vigente=0,MotivoRetiro=@Motivo,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ReferenciaId=@ReferenciaId;
        ELSE
            UPDATE dbo.ReferenciaComprobanteExterno SET ConfirmacionId=@ConfirmacionId,FechaDocumento=@FechaDocumento,CodigoTipoDocumento=@CodigoTipoDocumento,Serie=@Serie,NumeroDocumento=@NumeroDocumento,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ReferenciaId=@ReferenciaId;
        SELECT @Nuevo=(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.Vigente,referencia.ConfirmacionId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ReferenciaComprobanteExterno AS referencia WHERE referencia.ReferenciaId=@ReferenciaId;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@ReferenciaId),@Accion NVARCHAR(40)=CASE WHEN @Retirar=1 THEN N'RETIRAR_FACTURADO' WHEN @Anterior IS NULL THEN N'MARCAR_FACTURADO' ELSE N'CORREGIR_FACTURADO' END,
            @EstadoNuevo NVARCHAR(30)=CASE WHEN @Retirar=1 THEN N'RETIRADO' ELSE N'FACTURADO' END,@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ReferenciaComprobanteExterno',@Registro,@Accion,@EstadoAnterior,@EstadoNuevo,@FechaDocumento,@Motivo,@Detalle;
        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@OrdenCompraId;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,N'FACTURADO_GUARDAR',@OrdenCompraId,@ReferenciaId,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @ReferenciaId AS ReferenciaId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Dashboard_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Dashboard_Obtener
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@PaginaPendientes INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar indicadores o reportes.',1;
        IF @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51601,N'[HYT_VALIDACION] Revisa el rango de fechas.',1;
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),@VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END)
        FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        SELECT p.PresupuestoId,p.NumeroPresupuesto AS NumeroInterno,p.FechaRegistro,p.TipoAtencion,revision.RevisionId,revision.NumeroRevision,revision.ClienteNombre,revision.TiendaNombre,revision.Estado,revision.CodigoMoneda,revision.Total
        INTO #Presupuestos FROM dbo.Presupuesto AS p CROSS APPLY(SELECT TOP(1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        WHERE @VerPresupuesto=1 AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        SELECT oc.OrdenCompraId,oc.PresupuestoId,oc.NumeroRegistroInterno,oc.NumeroOCCliente,oc.FechaRegistro,oc.Estado,revision.CodigoMoneda,revision.Total,revision.ClienteNombre,revision.TiendaNombre,p.NumeroPresupuesto AS PresupuestoNumero,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.Estado AS EstadoEjecucion,confirmacion.ConfirmacionId,referencia.ReferenciaId
        INTO #Ordenes FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=oc.RevisionId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=oc.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=oc.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        LEFT JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1
        WHERE @VerOC=1 AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        IF @PaginaPendientes<1 OR @PaginaPendientes>100000000 THROW 51601,N'[HYT_VALIDACION] Página no válida.',1;
        DECLARE @VerPendientes BIT=0;
        IF @VerPresupuesto=1 AND @VerOC=1 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permiso.CodigoPermiso=N'GUIAS.CONSULTAR') SET @VerPendientes=1;
        SELECT p.PresupuestoId AS Id,p.NumeroPresupuesto AS Numero,p.FechaRegistro,p.TipoAtencion,revision.ClienteNombre,revision.TiendaNombre,
            CONVERT(BIT,CASE WHEN oc.OrdenCompraId IS NULL THEN 1 ELSE 0 END) AS SinOC,CONVERT(BIT,CASE WHEN guia.GuiaId IS NULL THEN 1 ELSE 0 END) AS SinGuia,
            CONVERT(BIT,CASE WHEN p.TipoAtencion=N'EMERGENCIA' AND revision.Estado<>N'APROBADO' THEN 1 ELSE 0 END) AS EmergenciaPorRegularizar,
            DATEDIFF(DAY,CONVERT(DATE,DATEADD(HOUR,-5,p.FechaRegistro)),CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()))) AS DiasPendiente
        INTO #Pendientes FROM dbo.Presupuesto AS p
        CROSS APPLY(SELECT TOP(1) r.ClienteNombre,r.TiendaNombre,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId AND oc.Estado<>N'CANCELADA'
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        WHERE @VerPendientes=1 AND revision.Estado<>N'CANCELADO' AND p.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (oc.OrdenCompraId IS NULL OR guia.GuiaId IS NULL OR (p.TipoAtencion=N'EMERGENCIA' AND revision.Estado<>N'APROBADO'));
        SELECT evento.FechaRegistro,evento.Tipo INTO #Eventos
        FROM (SELECT confirmacion.OrdenCompraId,confirmacion.FechaRegistro,N'CONFIRMACION' AS Tipo FROM dbo.ConfirmacionFacturacion AS confirmacion WHERE confirmacion.Vigente=1
            UNION ALL SELECT referencia.OrdenCompraId,referencia.FechaRegistro,N'FACTURADO' FROM dbo.ReferenciaComprobanteExterno AS referencia WHERE referencia.Vigente=1) AS evento
        INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=evento.OrdenCompraId
        INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=oc.RevisionId
        WHERE @VerOC=1 AND oc.Estado<>N'CANCELADA' AND evento.FechaRegistro>=@Desde AND evento.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%');
        SELECT (SELECT @VerPresupuesto AS VerPresupuestos,@VerOC AS VerOC,@VerPendientes AS VerPendientes,SYSUTCDATETIME() AS ActualizadoUTC,@Desde AS DesdeUTC,@HastaExclusiva AS HastaExclusivaUTC,
            (SELECT COUNT_BIG(1) FROM #Pendientes) AS TotalPendientes,
            (SELECT COUNT_BIG(1) FROM #Pendientes AS pendiente WHERE pendiente.FechaRegistro>=@Desde) AS PendientesPeriodo,
            JSON_QUERY(COALESCE((SELECT pendiente.Id,pendiente.Numero,pendiente.FechaRegistro,pendiente.TipoAtencion,pendiente.ClienteNombre,pendiente.TiendaNombre,pendiente.SinOC,pendiente.SinGuia,pendiente.EmergenciaPorRegularizar,pendiente.DiasPendiente FROM #Pendientes AS pendiente ORDER BY pendiente.FechaRegistro,pendiente.Id OFFSET (CONVERT(BIGINT,@PaginaPendientes)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH),N'[]')) AS Pendientes,
            JSON_QUERY(COALESCE((SELECT presupuesto.Estado,COUNT_BIG(1) AS Cantidad FROM #Presupuestos AS presupuesto GROUP BY presupuesto.Estado FOR JSON PATH),N'[]')) AS EstadosPresupuesto,
            JSON_QUERY(COALESCE((SELECT orden.Estado,COUNT_BIG(1) AS Cantidad FROM #Ordenes AS orden GROUP BY orden.Estado FOR JSON PATH),N'[]')) AS EstadosOC,
            (SELECT COUNT_BIG(1) FROM #Presupuestos) AS Presupuestos,
            (SELECT COUNT_BIG(1) FROM #Ordenes) AS Ordenes,
            (SELECT COUNT_BIG(1) FROM #Presupuestos AS presupuesto WHERE NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=presupuesto.PresupuestoId)) AS PresupuestosSinOC,
            (SELECT COUNT_BIG(1) FROM #Presupuestos AS presupuesto WHERE NOT EXISTS(SELECT 1 FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=presupuesto.PresupuestoId)) AS PresupuestosSinGuia,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NULL) AS PendientesFacturar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ConfirmacionId IS NULL) AS SinConfirmar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ConfirmacionId IS NOT NULL AND orden.ReferenciaId IS NULL) AS ConfirmadasSinMarcar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NULL AND orden.ConfirmacionId IS NOT NULL AND (orden.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') OR orden.EstadoEjecucion=N'TERMINADA') AND orden.GuiaId IS NOT NULL) AS ListasParaMarcar,
            (SELECT COUNT_BIG(1) FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' AND orden.ReferenciaId IS NOT NULL) AS Facturadas,
            JSON_QUERY(COALESCE((SELECT presupuesto.CodigoMoneda,SUM(presupuesto.Total) AS Importe,COUNT_BIG(1) AS Cantidad,SUM(CONVERT(BIGINT,CASE WHEN presupuesto.Total IS NULL THEN 1 ELSE 0 END)) AS PreciosPendientes FROM #Presupuestos AS presupuesto WHERE presupuesto.Estado<>N'CANCELADO' GROUP BY presupuesto.CodigoMoneda FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS ImportesPresupuesto,
            JSON_QUERY(COALESCE((SELECT orden.CodigoMoneda,SUM(orden.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Ordenes AS orden WHERE orden.Estado<>N'CANCELADA' GROUP BY orden.CodigoMoneda FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS ImportesOC,
            JSON_QUERY(COALESCE((SELECT movimiento.Fecha,SUM(movimiento.Presupuestos) AS Presupuestos,SUM(movimiento.Ordenes) AS Ordenes,SUM(movimiento.Confirmaciones) AS Confirmaciones,SUM(movimiento.Marcados) AS Marcados FROM
                (SELECT CONVERT(DATE,DATEADD(HOUR,-5,presupuesto.FechaRegistro)) AS Fecha,CONVERT(BIGINT,1) AS Presupuestos,CONVERT(BIGINT,0) AS Ordenes,CONVERT(BIGINT,0) AS Confirmaciones,CONVERT(BIGINT,0) AS Marcados FROM #Presupuestos AS presupuesto
                 UNION ALL SELECT CONVERT(DATE,DATEADD(HOUR,-5,orden.FechaRegistro)),CONVERT(BIGINT,0),CONVERT(BIGINT,1),CONVERT(BIGINT,0),CONVERT(BIGINT,0) FROM #Ordenes AS orden
                 UNION ALL SELECT CONVERT(DATE,DATEADD(HOUR,-5,evento.FechaRegistro)),CONVERT(BIGINT,0),CONVERT(BIGINT,0),CONVERT(BIGINT,CASE WHEN evento.Tipo=N'CONFIRMACION' THEN 1 ELSE 0 END),CONVERT(BIGINT,CASE WHEN evento.Tipo=N'FACTURADO' THEN 1 ELSE 0 END) FROM #Eventos AS evento) AS movimiento
                GROUP BY movimiento.Fecha ORDER BY movimiento.Fecha FOR JSON PATH),N'[]')) AS Movimientos
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) AS Datos;
        DROP TABLE #Presupuestos;DROP TABLE #Ordenes;DROP TABLE #Pendientes;DROP TABLE #Eventos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_Reporte_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Reporte_Listar
    @UsuarioId NVARCHAR(128),@Reporte NVARCHAR(40),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,
    @Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@Situacion NVARCHAR(30)=NULL,
    @Estado NVARCHAR(40)=NULL,@TipoAtencion NVARCHAR(20)=NULL,@Busqueda NVARCHAR(200)=NULL,
    @PendienteGuia BIT=0,@PendienteOC BIT=0,@PendienteConfirmacion BIT=0,@PendienteFacturar BIT=0,@Exportar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @VerPresupuesto BIT=0,@VerOC BIT=0,@VerGuia BIT=0,@VerConfirmacion BIT=0,@VerFactura BIT=0,@VerEjecucion BIT=0;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'REPORTES.CONSULTAR')
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar reportes.',1;
        SELECT @VerPresupuesto=MAX(CASE WHEN permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR' THEN 1 ELSE 0 END),
            @VerOC=MAX(CASE WHEN permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR' THEN 1 ELSE 0 END),
            @VerGuia=MAX(CASE WHEN permiso.CodigoPermiso=N'GUIAS.CONSULTAR' THEN 1 ELSE 0 END),
            @VerConfirmacion=MAX(CASE WHEN permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR' THEN 1 ELSE 0 END),
            @VerFactura=MAX(CASE WHEN permiso.CodigoPermiso=N'FACTURADO.CONSULTAR' THEN 1 ELSE 0 END),
            @VerEjecucion=MAX(CASE WHEN permiso.CodigoPermiso=N'EJECUCION.CONSULTAR' THEN 1 ELSE 0 END)
        FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        DECLARE @Legado NVARCHAR(40)=@Reporte;
        SET @Reporte=CASE WHEN @Reporte=N'GANADOS' THEN N'PRESUPUESTO' WHEN @Reporte IN(N'OC_COMPLETADAS',N'OC_CON_ACTA',N'OC_PENDIENTES_FACTURA') THEN N'OC_CLIENTE' ELSE @Reporte END;
        IF @Reporte IS NULL OR @Reporte NOT IN(N'PRESUPUESTO',N'OC_CLIENTE') OR @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @Desde>=@HastaExclusiva
            THROW 51601,N'[HYT_VALIDACION] Revisa reporte, fechas y página.',1;
        IF (@Reporte=N'PRESUPUESTO' AND @VerPresupuesto=0) OR (@Reporte=N'OC_CLIENTE' AND @VerOC=0)
            OR (@PendienteGuia=1 AND @VerGuia=0) OR (@PendienteOC=1 AND @VerOC=0) OR (@PendienteConfirmacion=1 AND @VerConfirmacion=0) OR (@PendienteFacturar=1 AND @VerFactura=0)
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para los documentos del reporte.',1;
        IF @TipoAtencion IS NOT NULL AND @TipoAtencion NOT IN(N'INVITACION',N'EMERGENCIA') THROW 51601,N'[HYT_VALIDACION] Tipo de atención no válido.',1;
        IF @Estado IS NOT NULL AND ((@Reporte=N'PRESUPUESTO' AND @Estado NOT IN(N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO',N'RECHAZADO',N'CANCELADO')) OR (@Reporte=N'OC_CLIENTE' AND @Estado NOT IN(N'REGISTRADA',N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE',N'CANCELADA'))) THROW 51601,N'[HYT_VALIDACION] Estado no válido para el reporte.',1;
        IF @Reporte=N'PRESUPUESTO' AND @Situacion IS NOT NULL THROW 51601,N'[HYT_VALIDACION] El seguimiento corresponde a OC.',1;
        IF @Situacion IS NOT NULL AND @Situacion NOT IN(N'SIN_CONFIRMAR',N'CONFIRMADAS',N'LISTAS',N'FACTURADAS') THROW 51601,N'[HYT_VALIDACION] Situación no válida.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.PresupuestoId ELSE oc.OrdenCompraId END AS Id,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.NumeroPresupuesto ELSE CONVERT(NVARCHAR(100),oc.NumeroRegistroInterno) END AS Numero,
            p.NumeroPresupuesto AS PresupuestoNumero,p.FechaRegistro AS PresupuestoRegistro,p.TipoAtencion,
            actual.ClienteNombre,actual.TiendaNombre,actual.Estado AS EstadoPresupuesto,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.Estado ELSE oc.Estado END AS Estado,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN p.FechaRegistro ELSE oc.FechaRegistro END AS FechaRegistro,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.CodigoMoneda ELSE fuente.CodigoMoneda END AS CodigoMoneda,
            CASE WHEN @Reporte=N'PRESUPUESTO' THEN actual.Total ELSE COALESCE(oc.Importe,fuente.Total) END AS Total,
            CASE WHEN @VerOC=1 THEN oc.NumeroRegistroInterno END AS OCInterno,CASE WHEN @VerOC=1 THEN oc.NumeroOCCliente END AS NumeroCliente,CASE WHEN @VerOC=1 THEN oc.FechaRegistro END AS OCRegistro,
            CASE WHEN @VerOC=1 THEN oc.Estado END AS EstadoOC,
            CASE WHEN @VerGuia=1 THEN guia.NumeroActa END AS Guia,CASE WHEN @VerGuia=1 AND guia.GuiaId IS NOT NULL THEN N'Guía' END AS TipoGuia,
            CASE WHEN @VerGuia=1 THEN guia.FechaRegistro END AS GuiaRegistro,CASE WHEN @VerGuia=1 THEN guia.FechaActa END AS GuiaFecha,
            CASE WHEN @VerEjecucion=1 THEN ejecucion.EjecucionId END AS EjecucionId,CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaRegistro END AS EjecucionRegistro,
            CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaInicioReal END AS EjecucionInicio,CASE WHEN @VerEjecucion=1 THEN ejecucion.FechaTerminoReal END AS EjecucionTermino,
            CASE WHEN @VerConfirmacion=1 THEN confirmacion.NumeroConfirmacion END AS Confirmacion,CASE WHEN @VerConfirmacion=1 THEN confirmacion.FechaRegistro END AS ConfirmacionRegistro,
            CASE WHEN @VerConfirmacion=1 THEN confirmacion.FechaConfirmacion END AS ConfirmacionFecha,
            CASE WHEN @VerFactura=1 THEN referencia.CodigoTipoDocumento END AS TipoComprobante,CASE WHEN @VerFactura=1 THEN referencia.Serie END AS Serie,
            CASE WHEN @VerFactura=1 THEN referencia.NumeroDocumento END AS NumeroComprobante,CASE WHEN @VerFactura=1 THEN referencia.FechaRegistro END AS FacturadoRegistro,
            CASE WHEN @VerFactura=1 THEN referencia.FechaDocumento END AS ComprobanteFecha,
            CONVERT(BIT,CASE WHEN @VerFactura=1 AND referencia.ReferenciaId IS NOT NULL THEN 1 ELSE 0 END) AS Facturado,
            CASE WHEN referencia.ReferenciaId IS NOT NULL THEN N'FACTURADA' WHEN confirmacion.ConfirmacionId IS NULL THEN N'SIN CONFIRMAR'
                 WHEN oc.Estado NOT IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') AND (ejecucion.Estado IS NULL OR ejecucion.Estado<>N'TERMINADA') THEN N'TRABAJO PENDIENTE'
                 WHEN guia.GuiaId IS NULL THEN N'GUÍA PENDIENTE' ELSE N'LISTA PARA MARCAR' END AS Situacion
        INTO #Reporte
        FROM dbo.Presupuesto AS p
        CROSS APPLY(SELECT TOP(1) r.RevisionId,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS actual
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=oc.RevisionId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        LEFT JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1
        WHERE ((@Reporte=N'PRESUPUESTO' AND p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva) OR (@Reporte=N'OC_CLIENTE' AND oc.FechaRegistro>=@Desde AND oc.FechaRegistro<@HastaExclusiva))
            AND (@Estado IS NULL OR (@Reporte=N'PRESUPUESTO' AND actual.Estado=@Estado) OR (@Reporte=N'OC_CLIENTE' AND oc.Estado=@Estado))
            AND (@TipoAtencion IS NULL OR p.TipoAtencion=@TipoAtencion)
            AND (@Cliente IS NULL OR actual.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR actual.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR actual.ClienteNombre LIKE @Patron ESCAPE N'~' OR actual.TiendaNombre LIKE @Patron ESCAPE N'~' OR (@VerOC=1 AND oc.NumeroOCCliente LIKE @Patron ESCAPE N'~') OR (@VerGuia=1 AND guia.NumeroActa LIKE @Patron ESCAPE N'~') OR (@VerConfirmacion=1 AND confirmacion.NumeroConfirmacion LIKE @Patron ESCAPE N'~'))
            AND (@PendienteGuia=0 OR guia.GuiaId IS NULL) AND (@PendienteOC=0 OR oc.OrdenCompraId IS NULL)
            AND (@PendienteConfirmacion=0 OR (confirmacion.ConfirmacionId IS NULL AND oc.Estado<>N'CANCELADA')) AND (@PendienteFacturar=0 OR (referencia.ReferenciaId IS NULL AND oc.Estado<>N'CANCELADA'))
            AND (@Legado<>N'GANADOS' OR actual.Estado=N'APROBADO')
            AND (@Legado<>N'OC_COMPLETADAS' OR (oc.Estado<>N'CANCELADA' AND (oc.Estado IN(N'TRABAJO_COMPLETADO',N'CERRADA_OPERATIVAMENTE') OR ejecucion.Estado=N'TERMINADA')))
            AND (@Legado<>N'OC_CON_ACTA' OR (oc.Estado<>N'CANCELADA' AND guia.GuiaId IS NOT NULL))
            AND (@Legado<>N'OC_PENDIENTES_FACTURA' OR (oc.Estado<>N'CANCELADA' AND referencia.ReferenciaId IS NULL));
        DELETE reporte FROM #Reporte AS reporte WHERE @Situacion IS NOT NULL AND NOT
            ((@Situacion=N'SIN_CONFIRMAR' AND reporte.Situacion=N'SIN CONFIRMAR') OR (@Situacion=N'CONFIRMADAS' AND reporte.Situacion IN(N'TRABAJO PENDIENTE',N'GUÍA PENDIENTE',N'LISTA PARA MARCAR')) OR (@Situacion=N'LISTAS' AND reporte.Situacion=N'LISTA PARA MARCAR') OR (@Situacion=N'FACTURADAS' AND reporte.Situacion=N'FACTURADA'));
        SELECT COALESCE((SELECT reporte.Id,reporte.Numero,reporte.PresupuestoNumero,reporte.PresupuestoRegistro,reporte.TipoAtencion,reporte.ClienteNombre,reporte.TiendaNombre,reporte.EstadoPresupuesto,reporte.Estado,reporte.FechaRegistro,reporte.CodigoMoneda,reporte.Total,
            reporte.OCInterno,reporte.NumeroCliente,reporte.OCRegistro,reporte.EstadoOC,reporte.Guia,reporte.TipoGuia,reporte.GuiaRegistro,reporte.GuiaFecha,
            reporte.EjecucionId,reporte.EjecucionRegistro,reporte.EjecucionInicio,reporte.EjecucionTermino,reporte.Confirmacion,reporte.ConfirmacionRegistro,reporte.ConfirmacionFecha,
            reporte.TipoComprobante,reporte.Serie,reporte.NumeroComprobante,reporte.FacturadoRegistro,reporte.ComprobanteFecha,reporte.Facturado,
            CASE WHEN @VerFactura=1 AND @VerConfirmacion=1 THEN reporte.Situacion END AS Situacion
        FROM #Reporte AS reporte ORDER BY reporte.FechaRegistro DESC,reporte.Id DESC OFFSET (CASE WHEN @Exportar=1 THEN 0 ELSE (CONVERT(BIGINT,@Pagina)-1)*20 END) ROWS FETCH NEXT (CASE WHEN @Exportar=1 THEN 2147483647 ELSE 20 END) ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM #Reporte;
        SELECT reporte.CodigoMoneda,SUM(reporte.Total) AS Importe,COUNT_BIG(1) AS Cantidad FROM #Reporte AS reporte GROUP BY reporte.CodigoMoneda ORDER BY reporte.CodigoMoneda;
        DROP TABLE #Reporte;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

