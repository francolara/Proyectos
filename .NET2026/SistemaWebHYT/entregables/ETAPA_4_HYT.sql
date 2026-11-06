-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Entrega incremental HYT etapa 4, para ejecución manual.
-- =============================================
-- Aplicar después de etapa 3. No ejecutado automáticamente.
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Códigos internos automáticos y precios gestionados desde materiales.

-- Fuente: Script/023_TarifarioMaterialDirecto.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET QUOTED_IDENTIFIER ON;
IF COL_LENGTH(N'dbo.Tarifario',N'EsTarifarioMaterial') IS NULL
    ALTER TABLE dbo.Tarifario ADD EsTarifarioMaterial BIT NOT NULL CONSTRAINT DF_Tarifario_Material DEFAULT(0) WITH VALUES;
GO
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID(N'dbo.Tarifario') AND name=N'UX_Tarifario_MaterialDirecto')
    CREATE UNIQUE INDEX UX_Tarifario_MaterialDirecto ON dbo.Tarifario(EsTarifarioMaterial) WHERE EsTarifarioMaterial=1;
GO


-- Fuente: StoreProcedure/spu_MantTarifario_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifario_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.TarifarioId AS Id, m.Codigo AS Codigo, m.Nombre AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            N'' AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.Tarifario AS m
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR m.Codigo LIKE @Patron ESCAPE N'~' OR m.Nombre LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1) AND m.EsTarifarioMaterial=0;
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifario_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifario_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.TarifarioId AS Id,
            m.Codigo,
            m.Nombre,
            m.ClienteId,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            CASE WHEN ref2.TipoEntidad=N'NATURAL' THEN CONCAT(ref2.Nombres, N' ', ref2.Apellidos) ELSE ref2.RazonSocial END AS ClienteIdLabel,
            CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS v WHERE v.TarifarioId=m.TarifarioId AND v.Estado<>N'BORRADOR') THEN 1 ELSE 0 END AS BIT) AS TieneVersionesPublicadas
        FROM dbo.Tarifario AS m LEFT JOIN dbo.Cliente AS ref2 ON ref2.ClienteId = m.ClienteId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.TarifarioId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'Tarifario' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'Tarifario' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifario_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifario_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @Codigo NVARCHAR(40);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Nombre' AND LEN(dato.value)>200) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Nombre.',1;
        DECLARE @Nombre NVARCHAR(200) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Nombre'))),N'');
        IF @Nombre IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Nombre.',1;
        DECLARE @ClienteId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId'));
        IF JSON_VALUE(@Datos,N'$.ClienteId') IS NOT NULL AND @ClienteId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: ClienteId.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.Codigo, anterior.Nombre, anterior.ClienteId, anterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.Tarifario AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.TarifarioId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @ClienteId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>@ClienteId) AND NOT EXISTS (SELECT 1 FROM dbo.Cliente AS destino WITH (HOLDLOCK) WHERE destino.ClienteId=@ClienteId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: ClienteId.',1;
        IF @Id IS NOT NULL AND COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.ClienteId')),-1)<>COALESCE(@ClienteId,-1) AND EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS v WHERE v.TarifarioId=@Id AND v.Estado<>N'BORRADOR') THROW 51218,N'[HYT_VALIDACION] No se puede reasignar el cliente de un tarifario publicado.',1;
        IF @Id IS NOT NULL SET @Codigo=JSON_VALUE(@Anterior,N'$.Codigo');
        ELSE
        BEGIN
            DECLARE @BloqueoCodigo INT,@SiguienteCodigo BIGINT,@TextoCodigo NVARCHAR(20);
            EXEC @BloqueoCodigo=sys.sp_getapplock @Resource=N'HYT:Codigo:Tarifario',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @BloqueoCodigo<0 THROW 51224,N'[HYT_VALIDACION] No se pudo asignar el código interno. Inténtalo nuevamente.',1;
            WHILE 1=1
            BEGIN
                SET @SiguienteCodigo=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                SET @TextoCodigo=CONVERT(NVARCHAR(20),@SiguienteCodigo);
                SET @Codigo=N'TAR-'+CASE WHEN LEN(@TextoCodigo)<6 THEN REPLICATE(N'0',6-LEN(@TextoCodigo)) ELSE N'' END+@TextoCodigo;
                IF NOT EXISTS(SELECT 1 FROM dbo.Tarifario AS existente WITH(UPDLOCK,HOLDLOCK) WHERE existente.Codigo=@Codigo) BREAK;
            END;
        END;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.Tarifario (Codigo, Nombre, ClienteId, Activo, UsuarioRegistro)
            VALUES (@Codigo, @Nombre, @ClienteId, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.Tarifario SET Codigo=@Codigo, Nombre=@Nombre, ClienteId=@ClienteId, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE TarifarioId=@Id;
        END;
        SELECT @Posterior=(SELECT posterior.Codigo, posterior.Nombre, posterior.ClienteId, posterior.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.Tarifario AS posterior WHERE posterior.TarifarioId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Tarifario',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
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


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.VersionId AS Id, CONVERT(NVARCHAR(40), m.NumeroVersion) AS Codigo, CONCAT(tarifa.Nombre, N' · Versión ', m.NumeroVersion) AS Descripcion, m.Estado AS Estado,
            m.Estado AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.TarifarioVersion AS m INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId = m.TarifarioId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR CONVERT(NVARCHAR(40), m.NumeroVersion) LIKE @Patron ESCAPE N'~' OR CONCAT(tarifa.Nombre, N' · Versión ', m.NumeroVersion) LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Estado = N'BORRADOR') AND tarifa.EsTarifarioMaterial=0 AND (@PadreId IS NULL OR m.TarifarioId = @PadreId);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.VersionId AS Id,
            m.TarifarioId,
            m.NumeroVersion,
            m.VigenciaDesde,
            m.VigenciaHasta,
            m.CodigoMonedaPredeterminada,
            m.CodigoTipoIGVPredeterminado,
            m.Estado,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            ref0.Nombre AS TarifarioIdLabel,
            JSON_QUERY((SELECT zona.ZonaId AS Id, catalogo.Codigo, catalogo.Descripcion FROM dbo.TarifarioVersionZona AS zona INNER JOIN dbo.Zona AS catalogo ON catalogo.Id=zona.ZonaId WHERE zona.VersionId = m.VersionId ORDER BY zona.ZonaId FOR JSON PATH)) AS Zonas
        FROM dbo.TarifarioVersion AS m INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId = m.TarifarioId LEFT JOIN dbo.Tarifario AS ref0 ON ref0.TarifarioId = m.TarifarioId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.VersionId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'TarifarioVersion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'TarifarioVersion' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @TarifarioId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TarifarioId'));
        IF @TarifarioId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: TarifarioId.',1;
        IF JSON_VALUE(@Datos,N'$.TarifarioId') IS NOT NULL AND @TarifarioId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: TarifarioId.',1;
        DECLARE @VigenciaDesde DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaDesde'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaDesde') IS NOT NULL AND @VigenciaDesde IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaDesde.',1;
        DECLARE @VigenciaHasta DATE = TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.VigenciaHasta'),23);
        IF JSON_VALUE(@Datos,N'$.VigenciaHasta') IS NOT NULL AND @VigenciaHasta IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VigenciaHasta.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMonedaPredeterminada' AND LEN(dato.value)>3) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMonedaPredeterminada.',1;
        DECLARE @CodigoMonedaPredeterminada NVARCHAR(3) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMonedaPredeterminada'))),N'');
        IF @CodigoMonedaPredeterminada IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoMonedaPredeterminada.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoTipoIGVPredeterminado' AND LEN(dato.value)>10) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoTipoIGVPredeterminado.',1;
        DECLARE @CodigoTipoIGVPredeterminado NVARCHAR(10) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoTipoIGVPredeterminado'))),N'');
        IF @CodigoTipoIGVPredeterminado IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoTipoIGVPredeterminado.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS padre WITH (UPDLOCK,HOLDLOCK) WHERE padre.TarifarioId=@TarifarioId) THROW 51208,N'[HYT_VALIDACION] El tarifario no existe.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.TarifarioId, anterior.NumeroVersion, anterior.VigenciaDesde, anterior.VigenciaHasta, anterior.CodigoMonedaPredeterminada, anterior.CodigoTipoIGVPredeterminado, anterior.Estado, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=anterior.VersionId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.TarifarioVersion AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.VersionId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TarifarioId'))<>@TarifarioId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @Id IS NOT NULL AND JSON_VALUE(@Anterior,N'$.Estado')<>N'BORRADOR' THROW 51211,N'[HYT_VALIDACION] La versión publicada es inmutable. Crea una nueva versión.',1;
        IF @TarifarioId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TarifarioId')),-1)<>@TarifarioId) AND NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS destino WITH (HOLDLOCK) WHERE destino.TarifarioId=@TarifarioId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: TarifarioId.',1;
        IF @VigenciaDesde IS NOT NULL AND @VigenciaHasta<@VigenciaDesde THROW 51219,N'[HYT_VALIDACION] Revisa las fechas de vigencia.',1;
        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
        DECLARE @NumeroVersion INT;
        SELECT @NumeroVersion=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.TarifarioVersion (TarifarioId, VigenciaDesde, VigenciaHasta, CodigoMonedaPredeterminada, CodigoTipoIGVPredeterminado, NumeroVersion, UsuarioRegistro)
            VALUES (@TarifarioId, @VigenciaDesde, @VigenciaHasta, @CodigoMonedaPredeterminada, @CodigoTipoIGVPredeterminado, @NumeroVersion, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.TarifarioVersion SET VigenciaDesde=@VigenciaDesde, VigenciaHasta=@VigenciaHasta, CodigoMonedaPredeterminada=@CodigoMonedaPredeterminada, CodigoTipoIGVPredeterminado=@CodigoTipoIGVPredeterminado,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@Id;
        END;
        DELETE FROM dbo.TarifarioVersionZona WHERE VersionId=@Id;
        INSERT INTO dbo.TarifarioVersionZona (VersionId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        SELECT @Posterior=(SELECT posterior.TarifarioId, posterior.NumeroVersion, posterior.VigenciaDesde, posterior.VigenciaHasta, posterior.CodigoMonedaPredeterminada, posterior.CodigoTipoIGVPredeterminado, posterior.Estado, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=posterior.VersionId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS posterior WHERE posterior.VersionId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
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


-- Fuente: StoreProcedure/spu_MantTarifarioPartida_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioPartida_Listar
    @UsuarioId NVARCHAR(128), @Pagina INT = 1, @Busqueda NVARCHAR(200) = NULL,
    @PadreId BIGINT = NULL, @SoloActivos BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina < 1 OR @Pagina > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT m.PartidaId AS Id, COALESCE(m.Posicion, CONVERT(NVARCHAR(40), m.PartidaId)) AS Codigo, material.Descripcion AS Descripcion, CASE WHEN m.Activo = 1 THEN N'ACTIVO' ELSE N'INACTIVO' END AS Estado,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX), z.Descripcion), N', ') FROM dbo.TarifarioPartidaZona AS pz INNER JOIN dbo.Zona AS z ON z.Id = pz.ZonaId WHERE pz.PartidaId = m.PartidaId) AS Informacion, m.FechaRegistro, creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion, editor.Nombres AS UsuarioActualizacion
        INTO #Pagina FROM dbo.TarifarioPartida AS m INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId = m.MaterialId INNER JOIN dbo.TarifarioVersion AS padre ON padre.VersionId = m.VersionId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id = m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id = m.UsuarioActualizacion
        WHERE (@Patron IS NULL OR COALESCE(m.Posicion, CONVERT(NVARCHAR(40), m.PartidaId)) LIKE @Patron ESCAPE N'~' OR material.Descripcion LIKE @Patron ESCAPE N'~' OR m.CodigoMaterialSAP LIKE @Patron ESCAPE N'~' OR material.CodigoInterno LIKE @Patron ESCAPE N'~') AND (@SoloActivos = 0 OR m.Activo = 1) AND (@PadreId IS NULL OR m.VersionId = @PadreId);
        SELECT p.Id, p.Codigo, p.Descripcion, p.Estado, p.Informacion,
            p.FechaRegistro, p.UsuarioRegistro, p.FechaActualizacion, p.UsuarioActualizacion
        FROM #Pagina AS p ORDER BY p.Codigo, p.Id
        OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS p;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioPartida_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioPartida_Obtener
    @UsuarioId NVARCHAR(128), @Id BIGINT, @PaginaAuditoria INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @PaginaAuditoria < 1 OR @PaginaAuditoria > 100000000 THROW 51201, N'[HYT_VALIDACION] Página no válida.', 1;
        SELECT (SELECT m.PartidaId AS Id,
            m.VersionId,
            m.MaterialId,
            m.Posicion,
            m.CodigoMaterialSAP,
            m.TipoServicioId,
            m.NumeroNecesidad,
            m.CodigoMoneda,
            m.CodigoTipoIGV,
            m.PrecioPEN,
            m.PrecioUSD,
            m.Activo,
            m.VersionFila,
            m.FechaRegistro,
            creador.Nombres AS UsuarioRegistro,
            m.FechaActualizacion,
            editor.Nombres AS UsuarioActualizacion,
            CONCAT(N'Versión ', ref0.NumeroVersion, N' / Tarifario ', ref0.TarifarioId) AS VersionIdLabel,
            ref1.Descripcion AS MaterialIdLabel,
            ref4.Descripcion AS TipoServicioIdLabel,
            JSON_QUERY((SELECT zona.ZonaId AS Id, catalogo.Codigo, catalogo.Descripcion FROM dbo.TarifarioPartidaZona AS zona INNER JOIN dbo.Zona AS catalogo ON catalogo.Id=zona.ZonaId WHERE zona.PartidaId = m.PartidaId ORDER BY zona.ZonaId FOR JSON PATH)) AS Zonas,
            padre.Estado AS VersionEstado,
            (SELECT cabecera.EsTarifarioMaterial FROM dbo.Tarifario AS cabecera WHERE cabecera.TarifarioId=padre.TarifarioId) AS EsTarifarioMaterial
        FROM dbo.TarifarioPartida AS m INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId = m.MaterialId INNER JOIN dbo.TarifarioVersion AS padre ON padre.VersionId = m.VersionId LEFT JOIN dbo.TarifarioVersion AS ref0 ON ref0.VersionId = m.VersionId LEFT JOIN dbo.ServicioMaterial AS ref1 ON ref1.MaterialId = m.MaterialId LEFT JOIN dbo.TipoServicio AS ref4 ON ref4.Id = m.TipoServicioId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=m.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=m.UsuarioActualizacion
        WHERE m.PartidaId=@Id FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion, usuario.Nombres AS Usuario, evento.FechaRegistro,
            evento.EstadoAnterior, evento.EstadoNuevo, evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'TarifarioPartida' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id)
        ORDER BY evento.FechaRegistro DESC, evento.EventoId DESC
        OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE evento.Entidad=N'TarifarioPartida' AND evento.RegistroId=CONVERT(NVARCHAR(100),@Id);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioPartida_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioPartida_Guardar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @Datos NVARCHAR(MAX),
    @Id BIGINT = NULL, @VersionOriginal BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT <> 0 THROW 51202, N'[HYT_VALIDACION] La operación administra su transacción.', 1;
        IF ISJSON(@Datos)<>1 THROW 51203, N'[HYT_VALIDACION] Datos no válidos.', 1;
        DECLARE @VersionId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.VersionId'));
        IF @VersionId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: VersionId.',1;
        IF JSON_VALUE(@Datos,N'$.VersionId') IS NOT NULL AND @VersionId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: VersionId.',1;
        DECLARE @MaterialId BIGINT = TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.MaterialId'));
        IF @MaterialId IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: MaterialId.',1;
        IF JSON_VALUE(@Datos,N'$.MaterialId') IS NOT NULL AND @MaterialId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: MaterialId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Posicion' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: Posicion.',1;
        DECLARE @Posicion NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Posicion'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMaterialSAP' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMaterialSAP.',1;
        DECLARE @CodigoMaterialSAP NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMaterialSAP'))),N'');
        DECLARE @TipoServicioId INT = TRY_CONVERT(INT,JSON_VALUE(@Datos,N'$.TipoServicioId'));
        IF JSON_VALUE(@Datos,N'$.TipoServicioId') IS NOT NULL AND @TipoServicioId IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: TipoServicioId.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'NumeroNecesidad' AND LEN(dato.value)>60) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: NumeroNecesidad.',1;
        DECLARE @NumeroNecesidad NVARCHAR(60) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.NumeroNecesidad'))),N'');
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoMoneda' AND LEN(dato.value)>3) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoMoneda.',1;
        DECLARE @CodigoMoneda NVARCHAR(3) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoMoneda'))),N'');
        IF @CodigoMoneda IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoMoneda.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'CodigoTipoIGV' AND LEN(dato.value)>10) THROW 51204,N'[HYT_VALIDACION] Texto demasiado largo: CodigoTipoIGV.',1;
        DECLARE @CodigoTipoIGV NVARCHAR(10) = NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoTipoIGV'))),N'');
        IF @CodigoTipoIGV IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: CodigoTipoIGV.',1;
        DECLARE @PrecioPEN DECIMAL(19,6) = TRY_CONVERT(DECIMAL(19,6),JSON_VALUE(@Datos,N'$.PrecioPEN'));
        IF JSON_VALUE(@Datos,N'$.PrecioPEN') IS NOT NULL AND @PrecioPEN IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: PrecioPEN.',1;
        IF @PrecioPEN<0 THROW 51207,N'[HYT_VALIDACION] El precio no puede ser negativo.',1;
        DECLARE @PrecioUSD DECIMAL(19,6) = TRY_CONVERT(DECIMAL(19,6),JSON_VALUE(@Datos,N'$.PrecioUSD'));
        IF JSON_VALUE(@Datos,N'$.PrecioUSD') IS NOT NULL AND @PrecioUSD IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: PrecioUSD.',1;
        IF @PrecioUSD<0 THROW 51207,N'[HYT_VALIDACION] El precio no puede ser negativo.',1;
        DECLARE @Activo BIT = TRY_CONVERT(BIT,CASE JSON_VALUE(@Datos,N'$.Activo') WHEN N'true' THEN N'1' WHEN N'false' THEN N'0' ELSE JSON_VALUE(@Datos,N'$.Activo') END);
        IF @Activo IS NULL THROW 51205,N'[HYT_VALIDACION] Dato obligatorio: Activo.',1;
        IF JSON_VALUE(@Datos,N'$.Activo') IS NOT NULL AND @Activo IS NULL THROW 51206,N'[HYT_VALIDACION] Formato no válido: Activo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX), @Posterior NVARCHAR(MAX), @EstadoAnterior NVARCHAR(40), @EstadoNuevo NVARCHAR(40);
        IF NOT EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS padre WITH (UPDLOCK,HOLDLOCK) WHERE padre.VersionId=@VersionId AND padre.Estado=N'BORRADOR') THROW 51208,N'[HYT_VALIDACION] Solo se pueden editar partidas de una versión borrador.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT anterior.VersionId, anterior.MaterialId, anterior.Posicion, anterior.CodigoMaterialSAP, anterior.TipoServicioId, anterior.NumeroNecesidad, anterior.CodigoMoneda, anterior.CodigoTipoIGV, anterior.PrecioPEN, anterior.PrecioUSD, anterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=anterior.PartidaId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.TarifarioPartida AS anterior WITH (UPDLOCK,HOLDLOCK)
            WHERE anterior.PartidaId=@Id AND anterior.VersionFila=@VersionOriginal;
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El registro cambió o ya no existe. Vuelve a consultarlo.',1;
        END;
        IF @Id IS NOT NULL AND TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.VersionId'))<>@VersionId THROW 51210,N'[HYT_VALIDACION] No se puede cambiar el registro padre existente.',1;
        IF @VersionId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.VersionId')),-1)<>@VersionId) AND NOT EXISTS (SELECT 1 FROM dbo.TarifarioVersion AS destino WITH (HOLDLOCK) WHERE destino.VersionId=@VersionId AND destino.Estado = N'BORRADOR') THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: VersionId.',1;
        IF @MaterialId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.MaterialId')),-1)<>@MaterialId) AND NOT EXISTS (SELECT 1 FROM dbo.ServicioMaterial AS destino WITH (HOLDLOCK) WHERE destino.MaterialId=@MaterialId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: MaterialId.',1;
        IF @TipoServicioId IS NOT NULL AND (@Id IS NULL OR COALESCE(TRY_CONVERT(BIGINT,JSON_VALUE(@Anterior,N'$.TipoServicioId')),-1)<>@TipoServicioId) AND NOT EXISTS (SELECT 1 FROM dbo.TipoServicio AS destino WITH (HOLDLOCK) WHERE destino.Id=@TipoServicioId AND destino.Activo = 1) THROW 51212,N'[HYT_VALIDACION] Referencia inexistente o no disponible: TipoServicioId.',1;
        IF JSON_QUERY(@Datos,N'$.ZonaIds') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.ZonaIds')),1)<>N'[' THROW 51220,N'[HYT_VALIDACION] Indica las zonas, incluso si la selección está vacía.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.ZonaIds') AS entrada LEFT JOIN dbo.Zona AS zona WITH (HOLDLOCK) ON zona.Id=TRY_CONVERT(INT,entrada.value) WHERE entrada.type<>2 OR zona.Id IS NULL OR (zona.Activo=0 AND NOT EXISTS (SELECT 1 FROM OPENJSON(@Anterior,N'$.Zonas') WITH (Id INT '$.Id') AS anteriorZona WHERE anteriorZona.Id=zona.Id)))
            THROW 51221,N'[HYT_VALIDACION] Una zona seleccionada ya no está disponible.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.TarifarioPartida (VersionId, MaterialId, Posicion, CodigoMaterialSAP, TipoServicioId, NumeroNecesidad, CodigoMoneda, CodigoTipoIGV, PrecioPEN, PrecioUSD, Activo, UsuarioRegistro)
            VALUES (@VersionId, @MaterialId, @Posicion, @CodigoMaterialSAP, @TipoServicioId, @NumeroNecesidad, @CodigoMoneda, @CodigoTipoIGV, @PrecioPEN, @PrecioUSD, @Activo, @UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE
        BEGIN
            UPDATE dbo.TarifarioPartida SET MaterialId=@MaterialId, Posicion=@Posicion, CodigoMaterialSAP=@CodigoMaterialSAP, TipoServicioId=@TipoServicioId, NumeroNecesidad=@NumeroNecesidad, CodigoMoneda=@CodigoMoneda, CodigoTipoIGV=@CodigoTipoIGV, PrecioPEN=@PrecioPEN, PrecioUSD=@PrecioUSD, Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PartidaId=@Id;
        END;
        DELETE FROM dbo.TarifarioPartidaZona WHERE PartidaId=@Id;
        INSERT INTO dbo.TarifarioPartidaZona (PartidaId,ZonaId)
        SELECT @Id,CONVERT(INT,seleccion.value) FROM OPENJSON(@Datos,N'$.ZonaIds') AS seleccion GROUP BY CONVERT(INT,seleccion.value);
        UPDATE dbo.TarifarioVersion SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@VersionId;
        SELECT @Posterior=(SELECT posterior.VersionId, posterior.MaterialId, posterior.Posicion, posterior.CodigoMaterialSAP, posterior.TipoServicioId, posterior.NumeroNecesidad, posterior.CodigoMoneda, posterior.CodigoTipoIGV, posterior.PrecioPEN, posterior.PrecioUSD, posterior.Activo, JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=posterior.PartidaId FOR JSON PATH)) AS Zonas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioPartida AS posterior WHERE posterior.PartidaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');
        SET @EstadoNuevo=JSON_VALUE(@Posterior,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Posterior) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN COALESCE(@EstadoAnterior,N'')<>COALESCE(@EstadoNuevo,N'') THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @RegistroId NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioPartida',@RegistroId,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
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


-- Fuente: StoreProcedure/spu_MaterialTarifas_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MaterialTarifas_Listar
    @UsuarioId NVARCHAR(128),@MaterialId BIGINT,@Pagina INT=1,@Busqueda NVARCHAR(200)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'MATERIALES.CONSULTAR')
            OR NOT EXISTS(SELECT 1 FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId AND permiso.CodigoPermiso=N'TARIFARIOS.CONSULTAR')
            THROW 51200,N'[HYT_PERMISO] Necesita permiso de consulta de materiales y tarifarios.',1;
        IF @MaterialId<0 OR @Pagina<1 OR @Pagina>100000000 THROW 51201,N'[HYT_VALIDACION] Material o página no válido.',1;
        SELECT COALESCE((SELECT partida.PartidaId AS Id,partida.MaterialId,material.CodigoInterno AS MaterialCodigo,material.Descripcion AS MaterialDescripcion,version.VersionId,tarifa.Nombre AS Tarifario,version.NumeroVersion,version.Estado,
            partida.Posicion,partida.CodigoMaterialSAP,partida.CodigoMoneda,partida.CodigoTipoIGV,partida.PrecioPEN,partida.PrecioUSD,partida.Activo,CAST(CASE WHEN tarifa.EsTarifarioMaterial=1 OR version.Estado=N'BORRADOR' THEN 1 ELSE 0 END AS BIT) AS Editable,
            tipo.Descripcion AS TipoServicio,partida.FechaRegistro,creador.Nombres AS UsuarioRegistro,partida.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=partida.PartidaId) AS Zonas
            FROM dbo.TarifarioPartida AS partida INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=partida.VersionId
            INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId
            INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId LEFT JOIN dbo.TipoServicio AS tipo ON tipo.Id=partida.TipoServicioId
            INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=partida.UsuarioRegistro LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=partida.UsuarioActualizacion
            WHERE (@MaterialId=0 OR partida.MaterialId=@MaterialId) AND (@Busqueda IS NULL OR material.CodigoInterno LIKE N'%'+@Busqueda+N'%' OR material.Descripcion LIKE N'%'+@Busqueda+N'%' OR partida.CodigoMaterialSAP LIKE N'%'+@Busqueda+N'%') ORDER BY partida.FechaRegistro DESC,partida.PartidaId DESC
            OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]') AS Filas;
        SELECT COUNT_BIG(1) AS Total FROM dbo.TarifarioPartida AS partida INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId WHERE (@MaterialId=0 OR partida.MaterialId=@MaterialId) AND (@Busqueda IS NULL OR material.CodigoInterno LIKE N'%'+@Busqueda+N'%' OR material.Descripcion LIKE N'%'+@Busqueda+N'%' OR partida.CodigoMaterialSAP LIKE N'%'+@Busqueda+N'%');
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_MaterialTarifas_Listar TO hyt_app_executor;
GO


-- Fuente: StoreProcedure/spu_MaterialTarifa_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MaterialTarifa_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@MaterialId BIGINT,
    @CodigoTipoIGV NVARCHAR(10),@PrecioPEN DECIMAL(19,6),@PrecioUSD DECIMAL(19,6)=NULL,@Activo BIT=1,
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF EXISTS(SELECT permiso.Codigo FROM (VALUES(N'MATERIALES.CONSULTAR'),(N'TARIFARIOS.CONSULTAR'),(N'TARIFARIOS.EDITAR')) AS permiso(Codigo)
            WHERE NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
                INNER JOIN dbo.RolPermiso AS acceso ON acceso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND acceso.CodigoPermiso=permiso.Codigo))
            THROW 51200,N'[HYT_PERMISO] No tiene permiso para administrar el tarifario.',1;
        IF @CodigoTipoIGV IS NULL OR @CodigoTipoIGV NOT IN(N'CON_IGV',N'SIN_IGV') OR (@PrecioPEN IS NULL AND @PrecioUSD IS NULL) OR @PrecioPEN<0 OR @PrecioUSD<0 OR @Activo IS NULL
            THROW 51227,N'[HYT_VALIDACION] Revisa el tipo de IGV, estado e ingresa al menos un precio en soles o dólares.',1;
        IF @PrecioPEN<>ROUND(@PrecioPEN,2) OR @PrecioUSD<>ROUND(@PrecioUSD,2)
            THROW 51227,N'[HYT_VALIDACION] Los precios admiten hasta dos decimales.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT,@VersionId BIGINT,@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@EstadoAnterior NVARCHAR(40),@EstadoNuevo NVARCHAR(40),@CodigoSAP NVARCHAR(60);
        SELECT @CodigoSAP=material.CodigoMaterialSAP FROM dbo.ServicioMaterial AS material WITH(HOLDLOCK) WHERE material.MaterialId=@MaterialId;
        IF NOT EXISTS(SELECT 1 FROM dbo.ServicioMaterial AS material WHERE material.MaterialId=@MaterialId)
            THROW 51227,N'[HYT_VALIDACION] El material no existe.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT precio.CodigoTipoIGV,precio.PrecioPEN,precio.PrecioUSD,precio.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@VersionId=precio.VersionId
            FROM dbo.TarifarioPartida AS precio WITH(UPDLOCK,HOLDLOCK) INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=precio.VersionId
            INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
            WHERE precio.PartidaId=@Id AND precio.MaterialId=@MaterialId AND precio.VersionFila=@VersionOriginal
              AND (tarifa.EsTarifarioMaterial=1 OR version.Estado=N'BORRADOR');
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El tarifario cambió o no permite edición. Vuelve a consultarlo.',1;
            UPDATE dbo.TarifarioPartida SET CodigoTipoIGV=@CodigoTipoIGV,PrecioPEN=@PrecioPEN,PrecioUSD=@PrecioUSD,Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PartidaId=@Id;
        END
        ELSE
        BEGIN
            DECLARE @Bloqueo INT;
            EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:TarifarioMaterialDirecto',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @Bloqueo<0 THROW 51227,N'[HYT_VALIDACION] No se pudo guardar. Intenta nuevamente.',1;
            SELECT @TarifarioId=tarifa.TarifarioId FROM dbo.Tarifario AS tarifa WITH(UPDLOCK,HOLDLOCK) WHERE tarifa.EsTarifarioMaterial=1;
            IF @TarifarioId IS NULL
            BEGIN
                DECLARE @Codigo NVARCHAR(40),@Numero BIGINT;
                WHILE 1=1
                BEGIN
                    SET @Numero=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                    SET @Codigo=N'TAR-MAT-'+CONVERT(NVARCHAR(20),@Numero);
                    IF NOT EXISTS(SELECT 1 FROM dbo.Tarifario AS tarifa WHERE tarifa.Codigo=@Codigo) BREAK;
                END;
                INSERT INTO dbo.Tarifario(Codigo,Nombre,Activo,UsuarioRegistro,EsTarifarioMaterial) VALUES(@Codigo,N'Tarifario de materiales',1,@UsuarioId,1);
                SET @TarifarioId=CONVERT(BIGINT,SCOPE_IDENTITY());
                INSERT INTO dbo.TarifarioVersion(TarifarioId,NumeroVersion,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro)
                VALUES(@TarifarioId,1,N'PEN',N'SIN_IGV',N'VIGENTE',@UsuarioId);
            END;
            SELECT @VersionId=version.VersionId FROM dbo.TarifarioVersion AS version WHERE version.TarifarioId=@TarifarioId AND version.NumeroVersion=1;
            IF @VersionId IS NULL THROW 51227,N'[HYT_VALIDACION] No se pudo preparar el tarifario.',1;
            INSERT INTO dbo.TarifarioPartida(VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,CodigoMaterialSAP,Activo,UsuarioRegistro)
            VALUES(@VersionId,@MaterialId,N'PEN',@CodigoTipoIGV,@PrecioPEN,@PrecioUSD,@CodigoSAP,@Activo,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.TarifarioPartidaZona(PartidaId,ZonaId) SELECT @Id,zona.ZonaId FROM dbo.MaterialZona AS zona WHERE zona.MaterialId=@MaterialId;
        END;
        SELECT @Nuevo=(SELECT precio.CodigoTipoIGV,precio.PrecioPEN,precio.PrecioUSD,precio.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioPartida AS precio WHERE precio.PartidaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');SET @EstadoNuevo=JSON_VALUE(@Nuevo,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN @EstadoAnterior<>@EstadoNuevo THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioPartida',@Registro,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH;
END;
GO


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Publicar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Publicar
    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT, @Numero INT, @Estado NVARCHAR(20);
        SELECT @TarifarioId=v.TarifarioId FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        IF @TarifarioId IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión ya no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS t WITH (UPDLOCK,HOLDLOCK) WHERE t.TarifarioId=@TarifarioId AND t.Activo=1)
            THROW 51222,N'[HYT_VALIDACION] El tarifario no está activo.',1;
        SELECT @Estado=v.Estado FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK)
        WHERE v.VersionId=@Id AND v.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión cambió. Vuelve a consultarla.',1;
        IF @Estado<>N'BORRADOR' THROW 51223,N'[HYT_VALIDACION] Solo se publica una versión borrador.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK) WHERE p.VersionId=@Id AND p.Activo=1)
            THROW 51224,N'[HYT_VALIDACION] Agrega al menos una partida activa antes de publicar.',1;
        IF EXISTS (SELECT 1 FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK)
            INNER JOIN dbo.ServicioMaterial AS material WITH (HOLDLOCK) ON material.MaterialId=p.MaterialId
            INNER JOIN dbo.UnidadMedida AS unidad WITH (HOLDLOCK) ON unidad.Id=material.UnidadMedidaId
            WHERE p.VersionId=@Id AND p.Activo=1 AND (p.PrecioPEN IS NULL OR material.Activo=0 OR unidad.Activo=0))
            THROW 51225,N'[HYT_VALIDACION] Revisa precios PEN, materiales y unidades de las partidas activas.',1;
        DECLARE @Anterior NVARCHAR(MAX), @Nuevo NVARCHAR(MAX), @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        SELECT @Anterior=(SELECT v.Estado,v.NumeroVersion,v.VigenciaDesde,v.VigenciaHasta FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        UPDATE dbo.TarifarioVersion SET Estado=N'VIGENTE',FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE VersionId=@Id;
        SELECT @Nuevo=(SELECT v.Estado,v.NumeroVersion,v.VigenciaDesde,v.VigenciaHasta FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'PUBLICAR',N'BORRADOR',N'VIGENTE',NULL,NULL,@Cambios;
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


-- Fuente: StoreProcedure/spu_MantTarifarioVersion_Clonar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTarifarioVersion_Clonar
    @UsuarioId NVARCHAR(128), @Id BIGINT, @VersionOriginal BINARY(8), @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT, @Numero INT, @Estado NVARCHAR(20);
        SELECT @TarifarioId=v.TarifarioId FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        IF @TarifarioId IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión ya no existe.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Tarifario AS t WITH (UPDLOCK,HOLDLOCK) WHERE t.TarifarioId=@TarifarioId AND t.Activo=1)
            THROW 51222,N'[HYT_VALIDACION] El tarifario no está activo.',1;
        SELECT @Estado=v.Estado FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK)
        WHERE v.VersionId=@Id AND v.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51209,N'[HYT_CONCURRENCIA] La versión cambió. Vuelve a consultarla.',1;
        SELECT @Numero=COALESCE(MAX(v.NumeroVersion),0)+1 FROM dbo.TarifarioVersion AS v WITH (UPDLOCK,HOLDLOCK) WHERE v.TarifarioId=@TarifarioId;
        INSERT INTO dbo.TarifarioVersion (TarifarioId,NumeroVersion,VigenciaDesde,VigenciaHasta,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro)
        SELECT v.TarifarioId,@Numero,v.VigenciaDesde,v.VigenciaHasta,v.CodigoMonedaPredeterminada,v.CodigoTipoIGVPredeterminado,N'BORRADOR',@UsuarioId
        FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@Id;
        DECLARE @NuevaVersion BIGINT=CONVERT(BIGINT,SCOPE_IDENTITY());
        INSERT INTO dbo.TarifarioVersionZona (VersionId,ZonaId)
        SELECT @NuevaVersion,z.ZonaId FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=@Id;
        DECLARE @Mapa TABLE (AnteriorId BIGINT NOT NULL PRIMARY KEY,NuevoId BIGINT NOT NULL);
        -- MERGE de solo inserción permite relacionar el identificador fuente con el generado.
        MERGE dbo.TarifarioPartida AS destino
        USING (SELECT p.PartidaId,p.MaterialId,p.CodigoMoneda,p.CodigoTipoIGV,p.PrecioPEN,p.PrecioUSD,p.Posicion,p.CodigoMaterialSAP,p.TipoServicioId,p.NumeroNecesidad,p.Activo
               FROM dbo.TarifarioPartida AS p WITH (HOLDLOCK) WHERE p.VersionId=@Id) AS fuente ON 1=0
        WHEN NOT MATCHED THEN INSERT (VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,Posicion,CodigoMaterialSAP,TipoServicioId,NumeroNecesidad,Activo,UsuarioRegistro)
        VALUES (@NuevaVersion,fuente.MaterialId,fuente.CodigoMoneda,fuente.CodigoTipoIGV,fuente.PrecioPEN,fuente.PrecioUSD,fuente.Posicion,fuente.CodigoMaterialSAP,fuente.TipoServicioId,fuente.NumeroNecesidad,fuente.Activo,@UsuarioId)
        OUTPUT fuente.PartidaId,inserted.PartidaId INTO @Mapa (AnteriorId,NuevoId);
        INSERT INTO dbo.TarifarioPartidaZona (PartidaId,ZonaId)
        SELECT mapa.NuevoId,zona.ZonaId FROM @Mapa AS mapa INNER JOIN dbo.TarifarioPartidaZona AS zona ON zona.PartidaId=mapa.AnteriorId;
        INSERT INTO dbo.EventoAuditoria (CorrelacionId,Entidad,RegistroId,Accion,EstadoNuevo,UsuarioId,CambiosJson)
        SELECT @CorrelacionId,N'TarifarioPartida',CONVERT(NVARCHAR(100),mapa.NuevoId),N'COPIAR',CASE p.Activo WHEN 1 THEN N'true' ELSE N'false' END,@UsuarioId,
            (SELECT mapa.AnteriorId AS PartidaOrigen,p.VersionId,p.MaterialId,p.CodigoMoneda,p.CodigoTipoIGV,p.PrecioPEN,p.PrecioUSD,p.Posicion,p.CodigoMaterialSAP,p.TipoServicioId,p.NumeroNecesidad,p.Activo,
                JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioPartidaZona AS z WHERE z.PartidaId=p.PartidaId FOR JSON PATH)) AS Zonas
             FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM @Mapa AS mapa INNER JOIN dbo.TarifarioPartida AS p ON p.PartidaId=mapa.NuevoId;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT @Id AS VersionOrigen,@NuevaVersion AS VersionNueva,@Numero AS NumeroVersion,
            JSON_QUERY((SELECT v.CodigoMonedaPredeterminada,v.CodigoTipoIGVPredeterminado,v.VigenciaDesde,v.VigenciaHasta,
                JSON_QUERY((SELECT z.ZonaId AS Id FROM dbo.TarifarioVersionZona AS z WHERE z.VersionId=v.VersionId FOR JSON PATH)) AS Zonas
             FROM dbo.TarifarioVersion AS v WHERE v.VersionId=@NuevaVersion FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Valores,
            (SELECT COUNT_BIG(1) FROM @Mapa AS mapa) AS PartidasCopiadas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@NuevaVersion);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'COPIAR',NULL,N'BORRADOR',NULL,NULL,@Cambios;
        SET @Registro=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioVersion',@Registro,N'GENERAR_COPIA',@Estado,@Estado,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @NuevaVersion AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTasaIGV_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTasaIGV_Listar
    @UsuarioId NVARCHAR(128), @Busqueda NVARCHAR(200)=NULL, @Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CONFIGURACION.CONSULTAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51201,N'[HYT_VALIDACION] Página no válida.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT tasa.FechaDesde,tasa.FechaHasta,tasa.Valor,alta.FechaRegistro AS FechaRegistro,creador.Nombres AS UsuarioRegistro,
            cambio.FechaRegistro AS FechaActualizacion,editor.Nombres AS UsuarioActualizacion,alta.Motivo
        INTO #Pagina
        FROM dbo.TasaIGV AS tasa
        OUTER APPLY (SELECT TOP (1) evento.FechaRegistro,evento.UsuarioId,evento.Motivo FROM dbo.EventoAuditoria AS evento
            WHERE evento.Entidad=N'TasaIGV' AND evento.RegistroId=CONVERT(NVARCHAR(10),tasa.FechaDesde,23) AND evento.Accion=N'NUEVA_VIGENCIA'
            ORDER BY evento.EventoId) AS alta
        LEFT JOIN dbo.AspNetUsers AS creador ON creador.Id=alta.UsuarioId
        OUTER APPLY (SELECT TOP (1) evento.FechaRegistro,evento.UsuarioId FROM dbo.EventoAuditoria AS evento
            WHERE evento.Entidad=N'TasaIGV' AND evento.RegistroId=CONVERT(NVARCHAR(10),tasa.FechaDesde,23) AND evento.Accion=N'CERRAR_VIGENCIA'
            ORDER BY evento.EventoId DESC) AS cambio
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=cambio.UsuarioId
        WHERE @Busqueda IS NULL OR CONVERT(NVARCHAR(10),tasa.FechaDesde,103) LIKE @Patron ESCAPE N'~'
            OR CONVERT(NVARCHAR(30),tasa.Valor) LIKE @Patron ESCAPE N'~';
        SELECT p.FechaDesde,p.FechaHasta,p.Valor,p.FechaRegistro,p.UsuarioRegistro,p.FechaActualizacion,p.UsuarioActualizacion,p.Motivo
        FROM #Pagina AS p ORDER BY p.FechaDesde DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT (SELECT COUNT_BIG(1) FROM #Pagina AS p) AS Total,MAX(tasa.FechaDesde) AS UltimaFecha FROM dbo.TasaIGV AS tasa;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_MantTasaIGV_Agregar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTasaIGV_Agregar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @FechaDesde DATE, @Valor DECIMAL(9,6),
    @Motivo NVARCHAR(1000), @UltimaFechaEsperada DATE=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CONFIGURACION.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @FechaDesde IS NULL OR @Valor IS NULL OR @Valor<0 OR @Valor>1 OR NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL
            THROW 51226,N'[HYT_VALIDACION] Indica fecha, tasa entre cero y uno y motivo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @UltimaFecha DATE,@AnteriorValor DECIMAL(9,6),@AnteriorHasta DATE;
        SELECT @UltimaFecha=MAX(tasa.FechaDesde) FROM dbo.TasaIGV AS tasa WITH (UPDLOCK,HOLDLOCK);
        IF (@UltimaFecha<>@UltimaFechaEsperada) OR (@UltimaFecha IS NULL AND @UltimaFechaEsperada IS NOT NULL) OR (@UltimaFecha IS NOT NULL AND @UltimaFechaEsperada IS NULL)
            THROW 51209,N'[HYT_CONCURRENCIA] Las vigencias cambiaron. Actualiza el listado antes de registrar otra.',1;
        IF @UltimaFecha IS NOT NULL AND @FechaDesde<=@UltimaFecha THROW 51227,N'[HYT_VALIDACION] La nueva vigencia debe ser posterior a la última registrada.',1;
        SELECT @AnteriorValor=tasa.Valor,@AnteriorHasta=tasa.FechaHasta FROM dbo.TasaIGV AS tasa WHERE tasa.FechaDesde=@UltimaFecha;
        UPDATE dbo.TasaIGV SET FechaHasta=@FechaDesde WHERE FechaDesde=@UltimaFecha;
        INSERT INTO dbo.TasaIGV (FechaDesde,FechaHasta,Valor) VALUES (@FechaDesde,NULL,@Valor);
        DECLARE @Registro NVARCHAR(100),@Detalle NVARCHAR(MAX);
        IF @UltimaFecha IS NOT NULL
        BEGIN
            SET @Registro=CONVERT(NVARCHAR(10),@UltimaFecha,23);
            SET @Detalle=(SELECT @AnteriorValor AS Valor,@AnteriorHasta AS FechaHastaAnterior,@FechaDesde AS FechaHastaNueva FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TasaIGV',@Registro,N'CERRAR_VIGENCIA',NULL,NULL,NULL,@Motivo,@Detalle;
        END;
        SET @Registro=CONVERT(NVARCHAR(10),@FechaDesde,23);
        SET @Detalle=(SELECT @Valor AS Valor,@FechaDesde AS FechaDesde,@UltimaFecha AS VigenciaAnterior FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TasaIGV',@Registro,N'NUEVA_VIGENCIA',NULL,NULL,NULL,@Motivo,@Detalle;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: Script/013_PermisosEtapa4.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifario_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifario_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifario_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioPartida_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioPartida_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioPartida_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MaterialTarifas_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MaterialTarifa_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Publicar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTarifarioVersion_Clonar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTasaIGV_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_MantTasaIGV_Agregar TO hyt_app_executor;
GO

