-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Actualiza rutas de adjuntos por carpeta/número y firma, conservando claves históricas. No mueve objetos ni modifica registros.
-- =============================================
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
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


-- Fuente: StoreProcedure/spu_ConfiguracionHYT_FirmaGuardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConfiguracionHYT_FirmaGuardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Datos NVARCHAR(2000),@VersionOriginal BINARY(8)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS u INNER JOIN dbo.AspNetUserRoles AS r ON r.UserId=u.Id INNER JOIN dbo.RolPermiso AS p ON p.RoleId=r.RoleId WHERE u.Id=@UsuarioId AND u.Activo=1 AND p.CodigoPermiso=N'CONFIGURACION.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para subir la firma.',1;
        IF ISJSON(@Datos)<>1 OR @Datos IS NULL THROW 51303,N'[HYT_VALIDACION] Firma no válida.',1;
        DECLARE @Clave NVARCHAR(100)=JSON_VALUE(@Datos,N'$.ClaveAlmacenamiento'),@Tipo NVARCHAR(50)=JSON_VALUE(@Datos,N'$.TipoContenido'),@Hash NVARCHAR(100)=JSON_VALUE(@Datos,N'$.SHA256'),@Tamano BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TamanoBytes'));
        DECLARE @NombreClave NVARCHAR(100)=CASE WHEN LEFT(@Clave,16) COLLATE Latin1_General_100_BIN2=N'ConfigHYT/Firma/' THEN SUBSTRING(@Clave,17,100) ELSE @Clave END;
        IF @Clave IS NULL OR LEN(@NombreClave)<>36 OR LEFT(@NombreClave,32) COLLATE Latin1_General_100_BIN2 LIKE N'%[^a-f0-9]%' OR RIGHT(@Clave,4) NOT IN(N'.png',N'.jpg')
            OR @Tipo IS NULL OR @Tipo NOT IN(N'image/png',N'image/jpeg') OR (@Tipo=N'image/png' AND RIGHT(@Clave,4)<>N'.png') OR (@Tipo=N'image/jpeg' AND RIGHT(@Clave,4)<>N'.jpg')
            OR @Hash IS NULL OR LEN(@Hash)<>64 OR @Hash COLLATE Latin1_General_100_BIN2 LIKE N'%[^A-Fa-f0-9]%' OR @Tamano IS NULL OR @Tamano NOT BETWEEN 1 AND 2097152
            OR NULLIF(JSON_VALUE(@Datos,N'$.NombreOriginal'),N'') IS NULL
            THROW 51305,N'[HYT_VALIDACION] Adjunta una firma JPG o PNG válida de hasta 2 MB.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(2000);
        SELECT @Anterior=c.FirmaPresupuestoJson FROM dbo.ConfiguracionHYT AS c WITH(UPDLOCK,HOLDLOCK) WHERE c.Id=1;
        IF NOT EXISTS(SELECT 1 FROM dbo.ConfiguracionHYT AS c WHERE c.Id=1 AND c.VersionFila=@VersionOriginal)
            THROW 51309,N'[HYT_CONCURRENCIA] Guarda los datos de HYT o vuelve a abrir la configuración antes de subir la firma.',1;
        UPDATE dbo.ConfiguracionHYT SET FirmaPresupuestoJson=@Datos,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE Id=1;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Datos) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfiguracionHYT',N'1',N'ACTUALIZAR_FIRMA',NULL,NULL,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
IF DATABASE_PRINCIPAL_ID(N'hyt_app_executor') IS NOT NULL GRANT EXECUTE ON dbo.spu_ConfiguracionHYT_FirmaGuardar TO hyt_app_executor;
GO

