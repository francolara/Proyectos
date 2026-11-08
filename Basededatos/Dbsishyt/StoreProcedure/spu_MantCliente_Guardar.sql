-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda Cliente con permisos, concurrencia, referencias y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Código automático, identidad según persona y número obligatorio; tipo de documento oculto sin validación, conservando el histórico y Código SAP/enlaces múltiples auditados por cliente.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida permisos específicos de alta, edición y activación del maestro.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCliente_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'CLIENTES.CREAR' ELSE N'CLIENTES.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoSAP.',1;
        DECLARE @CodigoSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoSAP'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'TipoEntidad' AND LEN(dato.value)>20) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: TipoEntidad.',1;
        DECLARE @TipoEntidad NVARCHAR(20) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.TipoEntidad'))),N'');
        IF @TipoEntidad IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: TipoEntidad.',1;
        DECLARE @TipoIdentificacion NVARCHAR(20);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'NumeroIdentificacion' AND LEN(dato.value)>30) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: NumeroIdentificacion.',1;
        DECLARE @NumeroIdentificacion NVARCHAR(30) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroIdentificacion'))),N'');
        IF @NumeroIdentificacion IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: NumeroIdentificacion.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'RazonSocial' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: RazonSocial.',1;
        DECLARE @RazonSocial NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RazonSocial'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombres' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombres.',1;
        DECLARE @Nombres NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombres'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Apellidos' AND LEN(dato.value)>150) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Apellidos.',1;
        DECLARE @Apellidos NVARCHAR(150) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Apellidos'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'DireccionFiscal' AND LEN(dato.value)>300) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: DireccionFiscal.',1;
        DECLARE @DireccionFiscal NVARCHAR(300) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.DireccionFiscal'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoUbigeo' AND LEN(dato.value)>6) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoUbigeo.',1;
        DECLARE @CodigoUbigeo NVARCHAR(6) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        DECLARE @UrlCargaDocumentos NVARCHAR(2048);
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.CodigoSAP, anterior.TipoEntidad, anterior.TipoIdentificacion, anterior.NumeroIdentificacion, anterior.RazonSocial, anterior.Nombres, anterior.Apellidos, anterior.DireccionFiscal, anterior.CodigoUbigeo, anterior.UrlCargaDocumentos, anterior.Activo, JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=anterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Cliente AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.ClienteId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'CLIENTES.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @TipoEntidad NOT IN (N'JURIDICA',N'NATURAL') THROW 51213,N'[HYT_VALIDACION] Tipo de entidad no válido.',1;
        IF (@TipoEntidad=N'JURIDICA' AND @RazonSocial IS NULL) OR (@TipoEntidad=N'NATURAL' AND (@Nombres IS NULL OR @Apellidos IS NULL)) THROW 51214,N'[HYT_VALIDACION] Completa la identificación del cliente.',1;
        SET @TipoIdentificacion=JSON_VALUE(@Anterior,N'$.TipoIdentificacion');
        IF @TipoEntidad=N'JURIDICA' BEGIN SET @Nombres=NULL; SET @Apellidos=NULL; END;
        ELSE SET @RazonSocial=NULL;
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51228,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        SET @UrlCargaDocumentos=JSON_VALUE(@Anterior,N'$.UrlCargaDocumentos');
        DECLARE @Enlaces NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Enlaces');
        IF EXISTS(SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Enlaces') AND (@Enlaces IS NULL OR LEFT(LTRIM(@Enlaces),1)<>N'[')
            THROW 51226,N'[HYT_VALIDACION] La lista de enlaces no es válida.',1;
        DECLARE @ListaEnlaces TABLE(Id BIGINT NULL,Descripcion NVARCHAR(MAX),Url NVARCHAR(MAX),Orden INT);
        IF @Enlaces IS NOT NULL
        BEGIN
            IF (SELECT COUNT_BIG(1) FROM OPENJSON(@Enlaces))>100 OR EXISTS(SELECT 1 FROM OPENJSON(@Enlaces) AS jsonEnlace WHERE jsonEnlace.type<>5)
                THROW 51226,N'[HYT_VALIDACION] Revisa la lista de enlaces (máximo cien).',1;
            INSERT INTO @ListaEnlaces(Id,Descripcion,Url,Orden)
            SELECT datos.Id,NULLIF(LTRIM(RTRIM(datos.Descripcion)),N''),NULLIF(LTRIM(RTRIM(datos.Url)),N''),CONVERT(INT,entrada.[key])
            FROM OPENJSON(@Enlaces) AS entrada
            CROSS APPLY OPENJSON(entrada.value) WITH(Id BIGINT '$.Id',Descripcion NVARCHAR(MAX) '$.Descripcion',Url NVARCHAR(MAX) '$.Url') AS datos;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS enlace WHERE enlace.Descripcion IS NULL OR LEN(enlace.Descripcion)>200 OR enlace.Url IS NULL OR LEN(enlace.Url)>2048
                OR (enlace.Url NOT LIKE N'https://%' AND enlace.Url NOT LIKE N'http://%') OR CHARINDEX(NCHAR(10),enlace.Url)>0 OR CHARINDEX(NCHAR(13),enlace.Url)>0)
                THROW 51226,N'[HYT_VALIDACION] Completa descripción y enlace HTTP o HTTPS válidos.',1;
            IF EXISTS(SELECT enlace.Id FROM @ListaEnlaces AS enlace WHERE enlace.Id IS NOT NULL GROUP BY enlace.Id HAVING COUNT_BIG(1)>1)
                THROW 51226,N'[HYT_VALIDACION] Un enlace está repetido.',1;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NOT NULL AND NOT EXISTS
                (SELECT 1 FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.Id=entrada.Id AND enlace.ClienteId=@Id))
                THROW 51226,N'[HYT_VALIDACION] Un enlace no pertenece al cliente o ya no existe.',1;
        END;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Cliente',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'CLI-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Cliente AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Cliente (Codigo, CodigoSAP, TipoEntidad, TipoIdentificacion, NumeroIdentificacion, RazonSocial, Nombres, Apellidos, DireccionFiscal, CodigoUbigeo, UrlCargaDocumentos, Activo, UsuarioRegistro)
            VALUES (@Codigo, @CodigoSAP, @TipoEntidad, @TipoIdentificacion, @NumeroIdentificacion, @RazonSocial, @Nombres, @Apellidos, @DireccionFiscal, @CodigoUbigeo, @UrlCargaDocumentos, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Cliente SET Codigo=@Codigo, CodigoSAP=@CodigoSAP, TipoEntidad=@TipoEntidad, TipoIdentificacion=@TipoIdentificacion, NumeroIdentificacion=@NumeroIdentificacion, RazonSocial=@RazonSocial, Nombres=@Nombres, Apellidos=@Apellidos, DireccionFiscal=@DireccionFiscal, CodigoUbigeo=@CodigoUbigeo, UrlCargaDocumentos=@UrlCargaDocumentos, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ClienteId=@Id;
        END;
        IF @Enlaces IS NOT NULL
        BEGIN
            DELETE enlace FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=@Id
                AND NOT EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id=enlace.Id);
            UPDATE enlace SET Descripcion=entrada.Descripcion,Url=entrada.Url,Orden=entrada.Orden,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
            FROM dbo.ClienteEnlaceDocumento AS enlace INNER JOIN @ListaEnlaces AS entrada ON entrada.Id=enlace.Id WHERE enlace.ClienteId=@Id;
            INSERT INTO dbo.ClienteEnlaceDocumento(ClienteId,Descripcion,Url,Orden,UsuarioRegistro)
            SELECT @Id,entrada.Descripcion,entrada.Url,entrada.Orden,@UsuarioId FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NULL;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.CodigoSAP, posterior.TipoEntidad, posterior.TipoIdentificacion, posterior.NumeroIdentificacion, posterior.RazonSocial, posterior.Nombres, posterior.Apellidos, posterior.DireccionFiscal, posterior.CodigoUbigeo, posterior.UrlCargaDocumentos, posterior.Activo, JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=posterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Cliente AS posterior WHERE posterior.ClienteId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Cliente',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
