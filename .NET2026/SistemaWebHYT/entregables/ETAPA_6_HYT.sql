-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Paquete incremental HYT etapa 6; aplicar después de etapa 5.
-- =============================================
-- Las fuentes individuales están aplicadas en la base local de desarrollo.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración mensual de presupuesto en consultas, reportes e impresiones; preserva claves internas. Aceptación automática auditada por registro de OC o Guía; consulta OC incluye guía y adjuntos por permiso; carpetas R2 por tipo y número con claves históricas compatibles. Guía como documento sin estados. KPI por estado y filtros de pendientes en listados con paginación SQL. Confirmación completa OC registrada con auditoría; transición manual admite confirmación vigente.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Actualiza consultas y filtros del dashboard, navegación comercial y resultados vacíos según fuentes de esta etapa.
-- =============================================

-- Fuente: Tablas/DocumentoOperacion.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.DocumentoOperacion',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.DocumentoOperacion
    (
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_DocumentoOperacion PRIMARY KEY,
        TipoOperacion NVARCHAR(40) NOT NULL,
        PadreId BIGINT NOT NULL,
        RegistroId BIGINT NOT NULL,
        ArchivoId BIGINT NULL,
        Huella CHAR(64) NOT NULL,
        UsuarioId NVARCHAR(128) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_DocumentoOperacion_Fecha DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_DocumentoOperacion_Usuario FOREIGN KEY (UsuarioId) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_DocumentoOperacion_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO

-- Fuente: Script/016_SeguimientoDocumental.sql
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

-- Fuente: Funciones/fn_ArchivoSegmento.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_ArchivoSegmento(@Numero NVARCHAR(100))
RETURNS NVARCHAR(100)
AS
BEGIN
    DECLARE @Origen NVARCHAR(100)=LTRIM(RTRIM(@Numero)),@Resultado NVARCHAR(100)=N'',@Indice INT=1,@Caracter NCHAR(1);
    WHILE @Indice<=LEN(@Origen COLLATE Latin1_General_100_BIN2)
    BEGIN
        SET @Caracter=SUBSTRING(@Origen COLLATE Latin1_General_100_BIN2,@Indice,1);
        SET @Resultado=@Resultado+CASE WHEN (@Caracter COLLATE Latin1_General_100_BIN2 LIKE N'[A-Za-z0-9]' OR @Caracter IN(N'_',N'-')) THEN @Caracter ELSE N'_' END;
        SET @Indice=@Indice+1;
    END;
    RETURN @Resultado;
END;
GO

-- Fuente: StoreProcedure/spu_SeguimientoCONFIRMACION_CambiarEstado.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'CONFIRMACION_ESTADO';
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
        DECLARE @PresupuestoId BIGINT=(SELECT oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId),@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH(UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió. Vuelve a consultarla.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END,@Anterior=(SELECT documento.FechaConfirmacion,documento.NumeroConfirmacion,documento.RepresentanteCliente,documento.Observacion,documento.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ConfirmacionFacturacion AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.ConfirmacionId=@Id AND documento.OrdenCompraId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'VIGENTE' AND @NuevoEstado=N'RETIRADA')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        UPDATE dbo.ConfirmacionFacturacion SET Vigente=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ConfirmacionId=@Id;
        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@PadreId;
        SELECT @Nuevo=(SELECT actual.FechaConfirmacion,actual.NumeroConfirmacion,actual.RepresentanteCliente,actual.Observacion,actual.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfirmacionFacturacion AS actual WHERE actual.ConfirmacionId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfirmacionFacturacion',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
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

-- Fuente: StoreProcedure/spu_SeguimientoCONFIRMACION_Guardar.sql
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


-- Fuente: StoreProcedure/spu_SeguimientoCONFIRMACION_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT documento.ConfirmacionId AS Id,documento.OrdenCompraId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) AS NumeroInterno,CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) AS NumeroDocumento,CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END AS Estado,documento.FechaConfirmacion AS FechaDocumento,
            fuente.CodigoMoneda,CONVERT(DECIMAL(19,2),NULL) AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),oc.VersionFila) AS VersionPadre,(SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=documento.ConfirmacionId) AS Adjuntos,(SELECT MIN(archivo.ArchivoId) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=documento.ConfirmacionId) AS PrimerArchivoId,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,CONVERT(NVARCHAR(MAX),NULL) AS Facturado
        INTO #Pagina FROM dbo.ConfirmacionFacturacion AS documento INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=documento.OrdenCompraId
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=oc.RevisionId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.ConfirmacionId=documento.ConfirmacionId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva AND (@Estado IS NULL OR CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END=@Estado)
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) LIKE @Patron ESCAPE N'~' OR CONVERT(NVARCHAR(100),documento.NumeroConfirmacion) LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_SeguimientoCONFIRMACION_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
        IF @Id IS NOT NULL SELECT @PadreReal=documento.OrdenCompraId FROM dbo.ConfirmacionFacturacion AS documento WHERE documento.ConfirmacionId=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.ConfirmacionId FROM dbo.ConfirmacionFacturacion AS documento WHERE documento.OrdenCompraId=@PadreId AND documento.Vigente=1;
        DECLARE @PresupuestoId BIGINT;
        SELECT @PresupuestoId=oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId;
        SELECT (SELECT N'CONFIRMACION' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            oc.VersionFila AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT documento.ConfirmacionId, documento.OrdenCompraId, documento.FechaConfirmacion, documento.NumeroConfirmacion, documento.RepresentanteCliente, documento.Observacion, documento.Vigente, documento.FechaRegistro, documento.FechaActualizacion, documento.VersionFila,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END AS EstadoDocumento
                FROM dbo.ConfirmacionFacturacion AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.ConfirmacionId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
            JSON_QUERY((SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256,archivo.Retirado,archivo.FechaRegistro,usuario.Nombres AS UsuarioRegistro,archivo.VersionFila
                FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId
                INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=archivo.UsuarioRegistro WHERE vinculo.ConfirmacionId=@Id ORDER BY archivo.ArchivoId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Adjuntos
        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ConfirmacionFacturacion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'ConfirmacionFacturacion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
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

-- Fuente: StoreProcedure/spu_SeguimientoEJECUCION_Guardar.sql
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

-- Fuente: StoreProcedure/spu_SeguimientoEJECUCION_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoEJECUCION_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'EJECUCION.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT documento.EjecucionId AS Id,documento.PresupuestoId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            CONVERT(NVARCHAR(100),documento.EjecucionId) AS NumeroInterno,CONVERT(NVARCHAR(100),documento.EjecucionId) AS NumeroDocumento,documento.Estado AS Estado,documento.FechaProgramada AS FechaDocumento,
            fuente.CodigoMoneda,CONVERT(DECIMAL(19,2),NULL) AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),p.VersionFila) AS VersionPadre,CONVERT(BIGINT,0) AS Adjuntos,CONVERT(BIGINT,NULL) AS PrimerArchivoId,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,CONVERT(NVARCHAR(MAX),NULL) AS Facturado
        INTO #Pagina FROM dbo.Ejecucion AS documento
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=documento.PresupuestoId OUTER APPLY (SELECT TOP (1) r.ClienteNombre,r.TiendaNombre,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva AND (@Estado IS NULL OR documento.Estado=@Estado)
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.EjecucionId) LIKE @Patron ESCAPE N'~' OR CONVERT(NVARCHAR(100),documento.EjecucionId) LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_SeguimientoEJECUCION_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoEJECUCION_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'EJECUCION.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
        IF @Id IS NOT NULL SELECT @PadreReal=documento.PresupuestoId FROM dbo.Ejecucion AS documento WHERE documento.EjecucionId=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.EjecucionId FROM dbo.Ejecucion AS documento WHERE documento.PresupuestoId=@PadreId;
        DECLARE @PresupuestoId BIGINT;
        SET @PresupuestoId=@PadreId;
        SELECT (SELECT N'EJECUCION' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            p.VersionFila AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT documento.EjecucionId, documento.PresupuestoId, documento.Responsable, documento.FechaProgramada, documento.FechaInicioReal, documento.FechaTerminoReal, documento.Descripcion, documento.Estado, documento.FechaRegistro, documento.FechaActualizacion, documento.VersionFila,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,documento.Estado AS EstadoDocumento
                FROM dbo.Ejecucion AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.EjecucionId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
            JSON_QUERY(N'[]') AS Adjuntos
        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Ejecucion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'Ejecucion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
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

-- Fuente: StoreProcedure/spu_SeguimientoGUIA_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista Guia (Acta de conformidad) por fecha/búsqueda, adjuntos y confirmación; veinte por página.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Filtros cliente y tienda conservados desde dashboard.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Incluye la OC vinculada al imprimir guías, con permiso de consulta.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'GUIAS.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT documento.GuiaId AS Id,documento.PresupuestoId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            CONVERT(NVARCHAR(100),documento.NumeroActa) AS NumeroInterno,CONVERT(NVARCHAR(100),documento.NumeroActa) AS NumeroDocumento,documento.Estado AS Estado,documento.FechaActa AS FechaDocumento,
            fuente.CodigoMoneda,CONVERT(DECIMAL(19,2),NULL) AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),p.VersionFila) AS VersionPadre,(SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=documento.GuiaId) AS Adjuntos,(SELECT MIN(archivo.ArchivoId) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=documento.GuiaId) AS PrimerArchivoId,
            CASE WHEN EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rolOC INNER JOIN dbo.RolPermiso AS permisoOC ON permisoOC.RoleId=rolOC.RoleId
                WHERE rolOC.UserId=@UsuarioId AND permisoOC.CodigoPermiso=N'OC_CLIENTE.CONSULTAR') THEN oc.NumeroOCCliente END AS NumeroOCCliente,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,CONVERT(NVARCHAR(MAX),NULL) AS Facturado
        INTO #Pagina FROM dbo.GuiaConformidad AS documento
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=documento.PresupuestoId OUTER APPLY (SELECT TOP (1) r.ClienteNombre,r.TiendaNombre,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva AND (@Estado IS NULL OR documento.Estado=@Estado)
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.NumeroActa) LIKE @Patron ESCAPE N'~' OR CONVERT(NVARCHAR(100),documento.NumeroActa) LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.NumeroOCCliente,pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_SeguimientoGUIA_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoGUIA_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'GUIAS.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
        IF @Id IS NOT NULL SELECT @PadreReal=documento.PresupuestoId FROM dbo.GuiaConformidad AS documento WHERE documento.GuiaId=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.GuiaId FROM dbo.GuiaConformidad AS documento WHERE documento.PresupuestoId=@PadreId;
        DECLARE @PresupuestoId BIGINT;
        SET @PresupuestoId=@PadreId;
        SELECT (SELECT N'GUIA' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            p.VersionFila AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT documento.GuiaId, documento.PresupuestoId, documento.FechaActa, documento.NumeroActa, documento.Observacion, documento.FechaAceptacion, documento.RepresentanteCliente, documento.Estado, documento.FechaRegistro, documento.FechaActualizacion, documento.VersionFila,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,documento.Estado AS EstadoDocumento
                FROM dbo.GuiaConformidad AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.GuiaId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
            JSON_QUERY((SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256,archivo.Retirado,archivo.FechaRegistro,usuario.Nombres AS UsuarioRegistro,archivo.VersionFila
                FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId
                INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=archivo.UsuarioRegistro WHERE vinculo.GuiaId=@Id ORDER BY archivo.ArchivoId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Adjuntos
        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'GuiaConformidad' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'GuiaConformidad' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
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


-- Fuente: StoreProcedure/spu_SeguimientoOC_Guardar.sql
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
        IF @EstadoRevision IS NULL OR @TotalRevision IS NULL OR @EstadoRevision NOT IN (N'BORRADOR',N'ENVIADO',N'OBSERVADO',N'APROBADO')
            THROW 51414,N'[HYT_VALIDACION] El presupuesto debe estar vigente y tener sus importes completos para registrar la OC.',1;
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
                OR LEN(JSON_VALUE(@ArchivoJson,N'$.SHA256'))<>64 OR @SHA LIKE '%[^0-9a-fA-F]%' OR @SHA IS NULL OR @Clave IS NULL OR (@Clave COLLATE Latin1_General_100_BIN2<>LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension AND @Clave COLLATE Latin1_General_100_BIN2<>N'OCCliente/'+dbo.fn_ArchivoSegmento(@NumeroOCCliente)+N'/'+LOWER(REPLACE(CONVERT(NVARCHAR(36),@CorrelacionId),N'-',N''))+@Extension)
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
        IF @Anterior IS NULL AND @EstadoRevision<>N'APROBADO'
        BEGIN
            DECLARE @FechaAceptacionAutomatica DATE=CONVERT(DATE,DATEADD(HOUR,-5,SYSUTCDATETIME()));
            UPDATE dbo.PresupuestoRevision SET Estado=N'APROBADO',FechaAprobacion=@FechaAceptacionAutomatica,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionId;
            DECLARE @RegistroRevision NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),
                @MotivoAceptacion NVARCHAR(1000)=N'Aceptación automática por registro de OC.',
                @DetalleAceptacion NVARCHAR(MAX)=(SELECT @EstadoRevision AS EstadoAnterior,N'APROBADO' AS EstadoNuevo,
                    @FechaAceptacionAutomatica AS FechaEvento,N'OC' AS TipoDocumento,@Id AS DocumentoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@RegistroRevision,N'CAMBIAR_ESTADO',
                @EstadoRevision,N'APROBADO',NULL,@MotivoAceptacion,@DetalleAceptacion;
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

-- Fuente: StoreProcedure/spu_SeguimientoOC_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(40)=NULL,@Confirmada BIT=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@PendienteConfirmacion BIT=0,@PendienteFacturar BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PendienteFacturar=1 AND NOT EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permiso.CodigoPermiso=N'FACTURADO.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para consultar facturación pendiente.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde THROW 51420,N'[HYT_VALIDACION] Revisa rango de fechas y página.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT documento.OrdenCompraId AS Id,documento.PresupuestoId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,fuente.ClienteNombre,fuente.TiendaNombre,
            CONVERT(NVARCHAR(100),documento.NumeroRegistroInterno) AS NumeroInterno,documento.NumeroOCCliente AS NumeroDocumento,documento.Estado AS Estado,documento.FechaEmisionOC AS FechaDocumento,
            fuente.CodigoMoneda,documento.Importe AS Importe,
            documento.FechaRegistro,creador.Nombres AS UsuarioRegistro,documento.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,CONVERT(BINARY(8),documento.VersionFila) AS VersionFila,
            CONVERT(BINARY(8),p.VersionFila) AS VersionPadre,(SELECT COUNT_BIG(1) FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.OrdenCompraId=documento.OrdenCompraId) AS Adjuntos,(SELECT MIN(archivo.ArchivoId) FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.OrdenCompraId=documento.OrdenCompraId) AS PrimerArchivoId,
            guia.GuiaId,guia.NumeroActa AS NumeroGuia,
            (SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=guia.GuiaId) AS AdjuntosGuia,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            (SELECT COUNT_BIG(1) FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.ConfirmacionId=confirmacion.ConfirmacionId) AS AdjuntosConfirmacion,(SELECT referencia.FechaDocumento,referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.FechaRegistro,autor.Nombres AS UsuarioRegistro FROM dbo.ReferenciaComprobanteExterno AS referencia INNER JOIN dbo.AspNetUsers AS autor ON autor.Id=referencia.UsuarioRegistro WHERE referencia.OrdenCompraId=documento.OrdenCompraId AND referencia.Vigente=1 AND EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permisoFactura ON permisoFactura.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permisoFactura.CodigoPermiso=N'FACTURADO.CONSULTAR') FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Facturado
        INTO #Pagina FROM dbo.OrdenCompraCliente AS documento
        INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=documento.PresupuestoId INNER JOIN dbo.PresupuestoRevision AS fuente ON fuente.RevisionId=documento.RevisionId LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=documento.OrdenCompraId AND confirmacion.Vigente=1
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId AND EXISTS
            (SELECT 1 FROM dbo.AspNetUserRoles AS rolGuia INNER JOIN dbo.RolPermiso AS permisoGuia ON permisoGuia.RoleId=rolGuia.RoleId WHERE rolGuia.UserId=@UsuarioId AND permisoGuia.CodigoPermiso=N'GUIAS.CONSULTAR')
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion
        WHERE documento.FechaRegistro>=@Desde AND documento.FechaRegistro<@HastaExclusiva
            AND (@Cliente IS NULL OR fuente.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR fuente.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR CONVERT(NVARCHAR(100),documento.NumeroRegistroInterno) LIKE @Patron ESCAPE N'~' OR documento.NumeroOCCliente LIKE @Patron ESCAPE N'~' OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR fuente.ClienteNombre LIKE @Patron ESCAPE N'~' OR fuente.TiendaNombre LIKE @Patron ESCAPE N'~')
            AND (@Confirmada IS NULL OR (@Confirmada=1 AND confirmacion.ConfirmacionId IS NOT NULL) OR (@Confirmada=0 AND confirmacion.ConfirmacionId IS NULL))
            AND (@PendienteConfirmacion=0 OR confirmacion.ConfirmacionId IS NULL)
            AND (@PendienteFacturar=0 OR NOT EXISTS(SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS pendiente WHERE pendiente.OrdenCompraId=documento.OrdenCompraId AND pendiente.Vigente=1));
        SELECT (SELECT pagina.Id,pagina.PadreId,pagina.PresupuestoId,pagina.PresupuestoNumero,pagina.TipoAtencion,pagina.ClienteNombre,pagina.TiendaNombre,pagina.NumeroInterno,pagina.NumeroDocumento,
            pagina.Estado,pagina.FechaDocumento,pagina.CodigoMoneda,pagina.Importe,pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.VersionFila,pagina.VersionPadre,pagina.Adjuntos,pagina.PrimerArchivoId,
            pagina.GuiaId,pagina.NumeroGuia,pagina.AdjuntosGuia,pagina.ConfirmacionId,pagina.NumeroConfirmacion,pagina.FechaConfirmacion,pagina.AdjuntosConfirmacion,JSON_QUERY(pagina.Facturado) AS Facturado
            FROM #Pagina AS pagina WHERE (@Estado IS NULL OR pagina.Estado=@Estado) ORDER BY pagina.FechaRegistro DESC,pagina.Id DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina WHERE (@Estado IS NULL OR pagina.Estado=@Estado);
        SELECT pagina.Estado,COUNT_BIG(1) AS Cantidad FROM #Pagina AS pagina GROUP BY pagina.Estado;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_SeguimientoOC_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoOC_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT=NULL,@PadreId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'OC_CLIENTE.CONSULTAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para esta operación documental.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51420,N'[HYT_VALIDACION] Página no válida.',1;
        IF @Id IS NULL AND @PadreId IS NULL THROW 51420,N'[HYT_VALIDACION] Selecciona el expediente.',1;
        DECLARE @PadreReal BIGINT;
        IF @Id IS NOT NULL SELECT @PadreReal=documento.PresupuestoId FROM dbo.OrdenCompraCliente AS documento WHERE documento.OrdenCompraId=@Id;
        IF @Id IS NOT NULL AND (@PadreReal IS NULL OR (@PadreId IS NOT NULL AND @PadreId<>@PadreReal)) THROW 51420,N'[HYT_VALIDACION] El registro no pertenece al expediente.',1;
        IF @Id IS NOT NULL SET @PadreId=@PadreReal;
        ELSE SELECT @Id=documento.OrdenCompraId FROM dbo.OrdenCompraCliente AS documento WHERE documento.PresupuestoId=@PadreId;
        DECLARE @PresupuestoId BIGINT;
        SET @PresupuestoId=@PadreId;
        SELECT (SELECT N'OC' AS Tipo,@PadreId AS PadreId,p.PresupuestoId,p.NumeroPresupuesto AS PresupuestoNumero,p.TipoAtencion,p.ClienteId,p.TiendaId,
            fuente.RevisionId AS RevisionPresupuestoId,fuente.NumeroRevision,fuente.ClienteNombre,fuente.TiendaNombre,fuente.Estado AS EstadoPresupuesto,fuente.CodigoMoneda,fuente.Total AS TotalPresupuesto,
            p.VersionFila AS VersionPadre,oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,
            guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,ejecucion.Estado AS EstadoEjecucion,
            confirmacion.ConfirmacionId,confirmacion.NumeroConfirmacion,confirmacion.FechaConfirmacion,
            JSON_QUERY((SELECT documento.OrdenCompraId, documento.PresupuestoId, documento.RevisionId, documento.NumeroOCCliente, documento.FechaEmisionOC, documento.FechaRecepcionOC, documento.Importe, documento.Observacion, documento.NumeroRegistroInterno, documento.CodigoMoneda, documento.Estado, documento.FechaRegistro, documento.FechaActualizacion, documento.VersionFila,creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,documento.Estado AS EstadoDocumento
                FROM dbo.OrdenCompraCliente AS documento INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=documento.UsuarioRegistro
                LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=documento.UsuarioActualizacion WHERE documento.OrdenCompraId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Documento,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.CodigoMoneda,r.Total,r.Estado FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            JSON_QUERY((SELECT anterior.ConfirmacionId,anterior.NumeroConfirmacion,anterior.FechaConfirmacion,anterior.Vigente FROM dbo.ConfirmacionFacturacion AS anterior WHERE anterior.OrdenCompraId=oc.OrdenCompraId ORDER BY anterior.ConfirmacionId DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Confirmaciones,
            JSON_QUERY((SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256,archivo.Retirado,archivo.FechaRegistro,usuario.Nombres AS UsuarioRegistro,archivo.VersionFila
                FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId
                INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=archivo.UsuarioRegistro WHERE vinculo.OrdenCompraId=@Id ORDER BY archivo.ArchivoId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Adjuntos
        FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS fuente
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.ConfirmacionFacturacion AS confirmacion ON confirmacion.OrdenCompraId=oc.OrdenCompraId AND confirmacion.Vigente=1
        WHERE p.PresupuestoId=@PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'OrdenCompraCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'OrdenCompraCliente' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Seguimiento_ObtenerArchivo.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Seguimiento_ObtenerArchivo
    @UsuarioId NVARCHAR(128),@Tipo NVARCHAR(20),@ArchivoId BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Permiso NVARCHAR(100)=CASE @Tipo WHEN N'OC' THEN N'OC_CLIENTE.CONSULTAR' WHEN N'GUIA' THEN N'GUIAS.CONSULTAR' WHEN N'CONFIRMACION' THEN N'CONFIRMACIONES.CONSULTAR' END;
        IF @Permiso IS NULL OR NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=@Permiso)
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para consultar este archivo.',1;
        SELECT archivo.ArchivoId,archivo.NombreOriginal,archivo.ClaveAlmacenamiento,archivo.TipoContenido,archivo.TamanoBytes,archivo.SHA256
        FROM dbo.ArchivoAdjunto AS archivo WHERE archivo.ArchivoId=@ArchivoId AND archivo.Retirado=0
          AND ((@Tipo=N'OC' AND EXISTS(SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId))
            OR (@Tipo=N'GUIA' AND EXISTS(SELECT 1 FROM dbo.GuiaArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId))
            OR (@Tipo=N'CONFIRMACION' AND EXISTS(SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo WHERE vinculo.ArchivoId=archivo.ArchivoId)));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/017_PermisosEtapa6.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoCONFIRMACION_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoEJECUCION_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoGUIA_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_SeguimientoOC_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Seguimiento_ObtenerArchivo TO hyt_app_executor;
GO

