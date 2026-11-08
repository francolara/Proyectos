-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Paquete incremental HYT etapa 5; aplicar después de etapa 4.
-- =============================================
-- Las fuentes individuales están aplicadas en la base local de desarrollo.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración mensual de presupuesto en consultas, reportes e impresiones; preserva claves internas. Listado incluye EjecucionId para botones documentales registrados. Guía como documento sin estados. KPI por estado y filtros de pendientes en listados con paginación SQL.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Actualiza consultas y filtros del dashboard, navegación comercial y resultados vacíos según fuentes de esta etapa.
-- =============================================

-- Fuente: Tablas/PresupuestoCorrelativoMensual.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PresupuestoCorrelativoMensual',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PresupuestoCorrelativoMensual
    (
        Periodo CHAR(6) NOT NULL CONSTRAINT PK_PresupuestoCorrelativoMensual PRIMARY KEY,
        Ultimo BIGINT NOT NULL CONSTRAINT CK_PresupuestoCorrelativoMensual_Ultimo CHECK(Ultimo>=0)
    );
END;
GO

-- Fuente: Script/026_NumeroPresupuestoMensual.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.Presupuesto',N'NumeroPresupuesto') IS NULL ALTER TABLE dbo.Presupuesto ADD NumeroPresupuesto NVARCHAR(40) NULL;
GO
BEGIN TRY
    BEGIN TRANSACTION;
    DECLARE @Filas BIGINT;
    SELECT @Filas=COUNT_BIG(1) FROM dbo.Presupuesto WITH(TABLOCKX,HOLDLOCK);
    ;WITH Numerados AS
    (
        SELECT p.PresupuestoId,CONVERT(CHAR(6),CONVERT(DATETIME2,p.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112) AS Periodo,
            ROW_NUMBER() OVER(PARTITION BY CONVERT(CHAR(6),CONVERT(DATETIME2,p.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112) ORDER BY p.FechaRegistro,p.PresupuestoId) AS Secuencia
        FROM dbo.Presupuesto AS p
    )
    UPDATE p SET NumeroPresupuesto=CONCAT(n.Periodo,CASE WHEN n.Secuencia<1000000 THEN RIGHT(CONCAT(N'000000',n.Secuencia),6) ELSE CONVERT(NVARCHAR(20),n.Secuencia) END)
    FROM dbo.Presupuesto AS p INNER JOIN Numerados AS n ON n.PresupuestoId=p.PresupuestoId WHERE p.NumeroPresupuesto IS NULL;
    ;WITH Maximos AS(SELECT LEFT(p.NumeroPresupuesto,6) AS Periodo,MAX(TRY_CONVERT(BIGINT,SUBSTRING(p.NumeroPresupuesto,7,34))) AS Ultimo FROM dbo.Presupuesto AS p GROUP BY LEFT(p.NumeroPresupuesto,6))
    UPDATE contador SET Ultimo=maximo.Ultimo FROM dbo.PresupuestoCorrelativoMensual AS contador INNER JOIN Maximos AS maximo ON maximo.Periodo=contador.Periodo WHERE contador.Ultimo<maximo.Ultimo;
    INSERT INTO dbo.PresupuestoCorrelativoMensual(Periodo,Ultimo)
    SELECT LEFT(p.NumeroPresupuesto,6),MAX(TRY_CONVERT(BIGINT,SUBSTRING(p.NumeroPresupuesto,7,34))) FROM dbo.Presupuesto AS p
    WHERE NOT EXISTS(SELECT 1 FROM dbo.PresupuestoCorrelativoMensual AS c WHERE c.Periodo=LEFT(p.NumeroPresupuesto,6)) GROUP BY LEFT(p.NumeroPresupuesto,6);
    ALTER TABLE dbo.Presupuesto ALTER COLUMN NumeroPresupuesto NVARCHAR(40) NOT NULL;
    IF NOT EXISTS(SELECT 1 FROM sys.indexes AS i WHERE i.object_id=OBJECT_ID(N'dbo.Presupuesto') AND i.name=N'UX_Presupuesto_NumeroPresupuesto') CREATE UNIQUE INDEX UX_Presupuesto_NumeroPresupuesto ON dbo.Presupuesto(NumeroPresupuesto);
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
    DECLARE @Mensaje NVARCHAR(4000)=ERROR_MESSAGE(),@Severidad INT=ERROR_SEVERITY(),@Estado INT=ERROR_STATE();
    RAISERROR(@Mensaje,@Severidad,@Estado);
END CATCH;
GO

-- Fuente: Script/027_NombreTipoServicioPresupuesto.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'TipoServicio') IS NOT NULL ALTER TABLE dbo.PresupuestoPartida ALTER COLUMN TipoServicio NVARCHAR(150) NULL;
GO

-- Fuente: Script/014_PresupuestosDocumentos.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'ZonaTexto') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD ZonaTexto NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'EmpresaSnapshotJson') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD EmpresaSnapshotJson NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'PoliticaCalculo') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD PoliticaCalculo NVARCHAR(20) NOT NULL CONSTRAINT DF_Revision_Politica DEFAULT N'LINEA_2_V1' WITH VALUES;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'FechaAprobacion') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD FechaAprobacion DATE NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'DescripcionTrabajo') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD DescripcionTrabajo NVARCHAR(1500) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'SimboloMoneda') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD SimboloMoneda NVARCHAR(5) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'TiendaCodigo') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD TiendaCodigo NVARCHAR(40) NULL;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'ZonasTexto') IS NULL
    ALTER TABLE dbo.PresupuestoPartida ADD ZonasTexto NVARCHAR(1000) NULL;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'OrigenPrecio') IS NULL
    ALTER TABLE dbo.PresupuestoPartida ADD OrigenPrecio NVARCHAR(20) NOT NULL CONSTRAINT DF_Partida_Origen DEFAULT N'MANUAL' WITH VALUES;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'MotivoPrecioManual') IS NULL
    ALTER TABLE dbo.PresupuestoPartida ADD MotivoPrecioManual NVARCHAR(300) NULL;
IF COL_LENGTH(N'dbo.ImpresionPresupuesto',N'ContenidoHTML') IS NULL
    ALTER TABLE dbo.ImpresionPresupuesto ADD ContenidoHTML NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.ImpresionPresupuesto',N'SHA256') IS NULL
    ALTER TABLE dbo.ImpresionPresupuesto ADD SHA256 CHAR(64) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_Revision_EmpresaJSON')
    ALTER TABLE dbo.PresupuestoRevision WITH CHECK ADD CONSTRAINT CK_Revision_EmpresaJSON CHECK (EmpresaSnapshotJson IS NULL OR ISJSON(EmpresaSnapshotJson)=1);
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_Tarifa_DosDecimales')
    ALTER TABLE dbo.TarifarioPartida WITH CHECK ADD CONSTRAINT CK_Tarifa_DosDecimales CHECK
        ((PrecioPEN IS NULL OR PrecioPEN=ROUND(PrecioPEN,2)) AND (PrecioUSD IS NULL OR PrecioUSD=ROUND(PrecioUSD,2)));
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_Partida_DosDecimales')
    ALTER TABLE dbo.PresupuestoPartida WITH CHECK ADD CONSTRAINT CK_Partida_DosDecimales CHECK
        (Cantidad=ROUND(Cantidad,2) AND (PrecioOrigen IS NULL OR PrecioOrigen=ROUND(PrecioOrigen,2)) AND (PrecioUnitario IS NULL OR PrecioUnitario=ROUND(PrecioUnitario,2)));
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_GuiaPartida_DosDecimales')
    ALTER TABLE dbo.GuiaPartida WITH CHECK ADD CONSTRAINT CK_GuiaPartida_DosDecimales CHECK (Cantidad=ROUND(Cantidad,2));
GO

-- Fuente: StoreProcedure/spu_Presupuesto_AyudaTarifario.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_AyudaTarifario
    @UsuarioId NVARCHAR(128),@ClienteId BIGINT,@Fecha DATE,@CodigoMoneda NVARCHAR(3)=N'PEN',@Busqueda NVARCHAR(200)=NULL,@Pagina INT=1,@MaterialId BIGINT=NULL,@PartidaId BIGINT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Fecha IS NULL OR @CodigoMoneda NOT IN (N'PEN',N'USD') THROW 51301,N'[HYT_VALIDACION] Revisa página, fecha y moneda.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT partida.PartidaId AS PartidaTarifarioId,material.MaterialId,material.CodigoInterno,COALESCE(partida.Posicion,material.Posicion) AS Posicion,
            COALESCE(partida.CodigoMaterialSAP,material.CodigoMaterialSAP) AS CodigoMaterialSAP,material.Descripcion,unidad.Codigo AS Unidad,
            tarifa.Codigo AS CodigoTarifario,version.NumeroVersion,partida.CodigoTipoIGV,@CodigoMoneda AS CodigoMoneda,
            CONVERT(DECIMAL(19,2),CASE WHEN @CodigoMoneda=N'USD' THEN partida.PrecioUSD ELSE partida.PrecioPEN END) AS Precio,
            servicio.Descripcion AS TipoServicio,partida.NumeroNecesidad,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=partida.PartidaId) AS Zonas
        INTO #Pagina FROM dbo.TarifarioPartida AS partida
        INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=partida.VersionId
        INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
        INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId
        INNER JOIN dbo.UnidadMedida AS unidad ON unidad.Id=material.UnidadMedidaId
        LEFT JOIN dbo.TipoServicio AS servicio ON servicio.Id=partida.TipoServicioId
        WHERE (@MaterialId IS NULL OR material.MaterialId=@MaterialId) AND (@PartidaId IS NULL OR partida.PartidaId=@PartidaId) AND partida.Activo=1 AND tarifa.Activo=1 AND material.Activo=1 AND unidad.Activo=1 AND version.Estado=N'VIGENTE'
            AND (tarifa.ClienteId IS NULL OR tarifa.ClienteId=@ClienteId)
            AND (version.VigenciaDesde IS NULL OR version.VigenciaDesde<=@Fecha) AND (version.VigenciaHasta IS NULL OR version.VigenciaHasta>=@Fecha)
            AND (@Busqueda IS NULL OR material.CodigoInterno LIKE @Patron ESCAPE N'~' OR material.Descripcion LIKE @Patron ESCAPE N'~' OR partida.CodigoMaterialSAP LIKE @Patron ESCAPE N'~' OR partida.Posicion LIKE @Patron ESCAPE N'~');
        SELECT pagina.PartidaTarifarioId,pagina.MaterialId,pagina.CodigoInterno,pagina.Posicion,pagina.CodigoMaterialSAP,pagina.Descripcion,pagina.Unidad,pagina.CodigoTarifario,pagina.NumeroVersion,
            pagina.CodigoTipoIGV,pagina.CodigoMoneda,pagina.Precio,pagina.TipoServicio,pagina.NumeroNecesidad,pagina.Zonas
        FROM #Pagina AS pagina ORDER BY pagina.CodigoInterno,pagina.PartidaTarifarioId OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
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

-- Fuente: StoreProcedure/spu_Presupuesto_DatosIniciales.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_DatosIniciales
    @UsuarioId NVARCHAR(128),@Fecha DATE,@ClienteId BIGINT=NULL,@TiendaId BIGINT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar presupuestos.',1;
        IF @Fecha IS NULL THROW 51305,N'[HYT_VALIDACION] Selecciona la fecha de emisión.',1;
        IF @TiendaId IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.Tienda AS tienda WHERE tienda.TiendaId=@TiendaId AND tienda.ClienteId=@ClienteId)
            THROW 51312,N'[HYT_VALIDACION] La tienda no pertenece al cliente.',1;
        SELECT (SELECT
            JSON_QUERY((SELECT tasa.FechaDesde,tasa.Valor FROM dbo.TasaIGV AS tasa WHERE tasa.FechaDesde<=@Fecha AND (tasa.FechaHasta IS NULL OR tasa.FechaHasta>@Fecha) FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)) AS Tasa,
            JSON_QUERY((SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.ConsideracionesPresupuesto FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Empresa,
            JSON_QUERY((SELECT cliente.ClienteId,cliente.Codigo,CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END AS Nombre,cliente.NumeroIdentificacion,CONCAT_WS(N' / ',NULLIF(cliente.DireccionFiscal,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=cliente.CodigoUbigeo)) AS DireccionFiscal
                FROM dbo.Cliente AS cliente WHERE cliente.ClienteId=@ClienteId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Cliente,
            JSON_QUERY((SELECT tienda.TiendaId,tienda.Codigo,tienda.Nombre,CONCAT_WS(N' / ',NULLIF(tienda.Direccion,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=tienda.CodigoUbigeo)) AS Direccion,tienda.CentroSAPFICO,zona.Descripcion AS ZonaTexto
                FROM dbo.Tienda AS tienda LEFT JOIN dbo.Zona AS zona ON zona.Id=tienda.ZonaId WHERE tienda.TiendaId=@TiendaId AND tienda.ClienteId=@ClienteId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Tienda,
            JSON_QUERY((SELECT cliente.CodigoSAP AS CodigoSAPProveedor,
                (SELECT TOP (1) proveedor.ContratoMarco FROM dbo.ProveedorCliente AS proveedor
                 WHERE proveedor.ClienteId=cliente.ClienteId AND proveedor.Activo=1 AND (proveedor.VigenciaDesde IS NULL OR proveedor.VigenciaDesde<=@Fecha) AND (proveedor.VigenciaHasta IS NULL OR proveedor.VigenciaHasta>=@Fecha)
                 ORDER BY proveedor.VigenciaDesde DESC,proveedor.Id DESC) AS ContratoMarco
                FROM dbo.Cliente AS cliente WHERE cliente.ClienteId=@ClienteId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Proveedor
        FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Presupuesto_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Datos NVARCHAR(MAX),
    @PresupuestoId BIGINT=NULL,@RevisionId BIGINT=NULL,@VersionOriginal BINARY(8)=NULL,
    @VersionPresupuesto BINARY(8)=NULL,@NuevaRevision BIT=0
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
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para guardar presupuestos.',1;
        IF ISJSON(@Datos)<>1 OR LEFT(LTRIM(@Datos),1)<>N'{' THROW 51303,N'[HYT_VALIDACION] Datos no válidos.',1;
        IF (@PresupuestoId IS NULL AND (@RevisionId IS NOT NULL OR @NuevaRevision=1)) OR (@PresupuestoId IS NOT NULL AND @RevisionId IS NULL)
            THROW 51304,N'[HYT_VALIDACION] Revisa el presupuesto y revisión.',1;
        DECLARE @ClienteId BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.ClienteId')),
            @TiendaId BIGINT=TRY_CONVERT(BIGINT,JSON_VALUE(@Datos,N'$.TiendaId')),
            @Tipo NVARCHAR(20)=JSON_VALUE(@Datos,N'$.TipoAtencion'),
            @Descripcion NVARCHAR(1500)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.Descripcion'))),N''),
            @Emision DATE=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEmision'),23),
            @Moneda NVARCHAR(3)=JSON_VALUE(@Datos,N'$.Revision.CodigoMoneda'),
            @TipoIGV NVARCHAR(10)=JSON_VALUE(@Datos,N'$.Revision.CodigoTipoIGV'),
            @FechaSolicitud DATE=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.FechaSolicitud'),23);
        IF @ClienteId IS NULL OR @TiendaId IS NULL OR @Descripcion IS NULL OR @Emision IS NULL OR @Tipo IS NULL OR @Tipo NOT IN (N'INVITACION',N'EMERGENCIA')
            OR @Moneda IS NULL OR JSON_VALUE(@Datos,N'$.Revision.CodigoMoneda') NOT IN (N'PEN',N'USD') OR @TipoIGV IS NULL OR @TipoIGV NOT IN (N'CON_IGV',N'SIN_IGV')
            THROW 51305,N'[HYT_VALIDACION] Completa cliente, tienda, tipo, descripción, fecha, moneda y tratamiento IGV.',1;
        IF JSON_VALUE(@Datos,N'$.FechaSolicitud') IS NOT NULL AND @FechaSolicitud IS NULL
            THROW 51305,N'[HYT_VALIDACION] Fecha de solicitud no válida.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Descripcion' AND LEN(dato.value)>1500)
            THROW 51306,N'[HYT_VALIDACION] La descripción excede 1500 caracteres.',1;
        IF JSON_QUERY(@Datos,N'$.Partidas') IS NULL OR LEFT(LTRIM(JSON_QUERY(@Datos,N'$.Partidas')),1)<>N'['
            THROW 51307,N'[HYT_VALIDACION] Indica las partidas, incluso cuando el borrador está vacío.',1;
        IF (SELECT COUNT_BIG(1) FROM OPENJSON(@Datos,N'$.Partidas'))>500
            THROW 51307,N'[HYT_VALIDACION] El presupuesto admite hasta 500 partidas.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Partidas') AS elemento WHERE elemento.type<>5)
            THROW 51307,N'[HYT_VALIDACION] Una partida no tiene el formato esperado.',1;
        DECLARE @Campos TABLE (Nombre NVARCHAR(80) NOT NULL,Limite INT NOT NULL);
        INSERT INTO @Campos (Nombre,Limite) VALUES
            (N'NumeroCotizacionCliente',60),(N'Trimestre',40),(N'VigenciaTexto',100),(N'CodigoSAPProveedor',60),(N'ContratoMarco',60),
            (N'Especialidad',150),(N'Coordinador',150),(N'Encargado',150),(N'Firmante',150),(N'CargoFirmante',150),(N'ZonaTexto',200),(N'Consideraciones',20000);
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Revision') AS dato INNER JOIN @Campos AS campo ON campo.Nombre=dato.[key] COLLATE DATABASE_DEFAULT WHERE LEN(dato.value)>campo.Limite)
            THROW 51306,N'[HYT_VALIDACION] Un campo de cabecera excede su longitud permitida.',1;
        IF JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion') IS NOT NULL AND TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion'),23) IS NULL
            THROW 51305,N'[HYT_VALIDACION] Fecha de ejecución no válida.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX),@AnteriorRevision NVARCHAR(MAX),@AnteriorPartidas NVARCHAR(MAX),@Estado NVARCHAR(20),
            @MonedaAnterior NVARCHAR(3),@ClienteAnterior BIGINT,@TiendaAnterior BIGINT,@TipoAnterior NVARCHAR(20),@NumeroRevision INT,
            @RevisionOrigen BIGINT=@RevisionId,@Registro NVARCHAR(100),@Accion NVARCHAR(60),@EsNuevo BIT=CASE WHEN @PresupuestoId IS NULL THEN 1 ELSE 0 END;
        IF @EsNuevo=0
        BEGIN
            SELECT @Anterior=(SELECT p.ClienteId,p.TiendaId,p.TipoAtencion,p.Descripcion,p.FechaSolicitud FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),
                @ClienteAnterior=p.ClienteId,@TiendaAnterior=p.TiendaId,@TipoAnterior=p.TipoAtencion
            FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId AND p.VersionFila=@VersionPresupuesto;
            IF @Anterior IS NULL THROW 51309,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
            SELECT @Estado=r.Estado,@MonedaAnterior=r.CodigoMoneda,
                @AnteriorRevision=(SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.NumeroCotizacionCliente,r.ClienteNombre,r.ClienteDocumento,r.ClienteDireccion,r.TiendaNombre,r.TiendaDireccion,
                    r.FechaEjecucion,r.Trimestre,r.VigenciaTexto,r.CodigoSAPProveedor,r.ContratoMarco,r.Especialidad,r.Coordinador,r.Encargado,r.Consideraciones,r.Firmante,r.CargoFirmante,
                    r.CodigoMoneda,r.CodigoTipoIGV,r.FechaDesdeTasaIGV,r.TasaIGVAplicada,r.Subtotal,r.Impuesto,r.Total,r.ZonaTexto,r.EmpresaSnapshotJson,r.DescripcionTrabajo,r.SimboloMoneda,r.TiendaCodigo,r.Estado FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
            FROM dbo.PresupuestoRevision AS r WITH (UPDLOCK,HOLDLOCK)
            WHERE r.PresupuestoId=@PresupuestoId AND r.RevisionId=@RevisionId AND r.VersionFila=@VersionOriginal;
            IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió o no pertenece al presupuesto.',1;
            IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (HOLDLOCK) WHERE oc.PresupuestoId=@PresupuestoId)
                THROW 51310,N'[HYT_VALIDACION] El presupuesto vinculado a una OC no admite cambios ni nuevas revisiones.',1;
            IF @Estado=N'CANCELADO' OR (@NuevaRevision=0 AND @Estado NOT IN (N'BORRADOR',N'OBSERVADO')) OR (@NuevaRevision=1 AND @Estado=N'BORRADOR')
                THROW 51310,N'[HYT_VALIDACION] Edita solo el registrado u observado; para una revisión enviada crea una nueva revisión.',1;
            IF EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.NumeroRevision>(SELECT origen.NumeroRevision FROM dbo.PresupuestoRevision AS origen WHERE origen.RevisionId=@RevisionId))
                THROW 51310,N'[HYT_VALIDACION] Abre la revisión más reciente del presupuesto.',1;
            IF (@ClienteAnterior<>@ClienteId OR @TiendaAnterior<>@TiendaId OR @TipoAnterior<>@Tipo)
                AND (@NuevaRevision=1 OR EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.Estado<>N'BORRADOR') OR EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS g WHERE g.PresupuestoId=@PresupuestoId))
                THROW 51310,N'[HYT_VALIDACION] No se puede cambiar cliente, tienda o tipo cuando ya existe historial enviado o guía.',1;
            IF @MonedaAnterior<>@Moneda AND COALESCE(JSON_VALUE(@Datos,N'$.RevalorizarMoneda'),N'false')<>N'true'
                THROW 51311,N'[HYT_VALIDACION] Confirma la revalorización de todas las partidas al cambiar la moneda.',1;
            SELECT @AnteriorPartidas=(SELECT p.NumeroLinea,p.Grupo,p.MaterialId,p.PartidaTarifarioId,p.Posicion,p.CodigoInterno,p.CodigoMaterialSAP,p.CentroSAPFICO,p.NumeroNecesidad,p.TipoServicio,p.Descripcion,p.Unidad,
                p.Cantidad,p.PrecioOrigen,p.CodigoMonedaOrigen,p.CodigoTipoIGVOrigen,p.PrecioUnitario,p.BaseImponible,p.Impuesto,p.Total,p.ZonasTexto,p.OrigenPrecio,p.MotivoPrecioManual
                FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId=@RevisionId ORDER BY p.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES);
        END;
        IF NOT EXISTS (SELECT 1 FROM dbo.Cliente AS cliente WITH (HOLDLOCK) INNER JOIN dbo.Tienda AS tienda WITH (HOLDLOCK) ON tienda.ClienteId=cliente.ClienteId
            WHERE cliente.ClienteId=@ClienteId AND tienda.TiendaId=@TiendaId AND ((cliente.Activo=1 AND tienda.Activo=1) OR (@EsNuevo=0 AND @ClienteAnterior=@ClienteId AND @TiendaAnterior=@TiendaId)))
            THROW 51312,N'[HYT_VALIDACION] Revisa cliente y tienda; la tienda debe pertenecer al cliente.',1;
        DECLARE @DesdeTasa DATE,@Tasa DECIMAL(9,6);
        IF @EsNuevo=0 AND @NuevaRevision=0 AND CONVERT(DATE,JSON_VALUE(@AnteriorRevision,N'$.FechaEmision'),23)=@Emision
        BEGIN
            SET @DesdeTasa=CONVERT(DATE,JSON_VALUE(@AnteriorRevision,N'$.FechaDesdeTasaIGV'),23);
            SET @Tasa=CONVERT(DECIMAL(9,6),JSON_VALUE(@AnteriorRevision,N'$.TasaIGVAplicada'));
        END
        ELSE
            SELECT @DesdeTasa=tasa.FechaDesde,@Tasa=tasa.Valor FROM dbo.TasaIGV AS tasa WITH (HOLDLOCK) WHERE tasa.FechaDesde<=@Emision AND (tasa.FechaHasta IS NULL OR tasa.FechaHasta>@Emision);
        IF @Tasa IS NULL THROW 51313,N'[HYT_VALIDACION] No existe una tasa IGV configurada para la fecha de emisión.',1;
        CREATE TABLE #Lineas
        (
            NumeroLinea INT NOT NULL PRIMARY KEY,PartidaAnterior BIGINT NULL,Grupo NVARCHAR(200) NULL,MaterialId BIGINT NULL,PartidaTarifarioId BIGINT NULL,
            Posicion NVARCHAR(60) NULL,CodigoInterno NVARCHAR(60) NULL,CodigoMaterialSAP NVARCHAR(60) NULL,CentroSAPFICO NVARCHAR(50) NULL,NumeroNecesidad NVARCHAR(60) NULL,
            TipoServicio NVARCHAR(150) NULL,Descripcion NVARCHAR(1500) NULL,Unidad NVARCHAR(30) NULL,Cantidad DECIMAL(19,6) NOT NULL,PrecioOrigen DECIMAL(19,6) NULL,
            CodigoMonedaOrigen NVARCHAR(3) NULL,CodigoTipoIGVOrigen NVARCHAR(10) NULL,PrecioUnitario DECIMAL(19,6) NULL,BaseImponible DECIMAL(19,2) NULL,Impuesto DECIMAL(19,2) NULL,
            Total DECIMAL(19,2) NULL,ZonasTexto NVARCHAR(1000) NULL,OrigenPrecio NVARCHAR(20) NULL,MotivoPrecioManual NVARCHAR(300) NULL,Conservar BIT NOT NULL DEFAULT 0
        );
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Partidas') WITH (Cantidad NVARCHAR(MAX),PrecioManual NVARCHAR(MAX)) AS entrada
            WHERE TRY_CONVERT(DECIMAL(19,6),entrada.Cantidad) IS NULL OR TRY_CONVERT(DECIMAL(19,6),entrada.Cantidad)<=0
            OR entrada.Cantidad LIKE N'%[^0-9.]%' OR (CHARINDEX(N'.',entrada.Cantidad)>0 AND LEN(entrada.Cantidad)-CHARINDEX(N'.',entrada.Cantidad)>2)
            OR (entrada.PrecioManual IS NOT NULL AND (TRY_CONVERT(DECIMAL(19,6),entrada.PrecioManual) IS NULL OR TRY_CONVERT(DECIMAL(19,6),entrada.PrecioManual)<0
                OR entrada.PrecioManual LIKE N'%[^0-9.]%' OR (CHARINDEX(N'.',entrada.PrecioManual)>0 AND LEN(entrada.PrecioManual)-CHARINDEX(N'.',entrada.PrecioManual)>2))))
            THROW 51314,N'[HYT_VALIDACION] Revisa cantidades y precios: positivos, hasta dos decimales y dentro del rango permitido.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos,N'$.Partidas') AS item CROSS APPLY OPENJSON(item.value) AS dato
            INNER JOIN (VALUES (N'Grupo',200),(N'Posicion',60),(N'CodigoInterno',60),(N'CodigoMaterialSAP',60),(N'CentroSAPFICO',50),(N'NumeroNecesidad',60),(N'TipoServicio',150),
                (N'Descripcion',1500),(N'Unidad',30),(N'ZonasTexto',1000),(N'MotivoPrecioManual',300)) AS limite(Nombre,Largo) ON limite.Nombre=dato.[key] COLLATE DATABASE_DEFAULT WHERE LEN(dato.value)>limite.Largo)
            THROW 51306,N'[HYT_VALIDACION] Un campo de partida excede su longitud permitida.',1;
        INSERT INTO #Lineas (NumeroLinea,PartidaAnterior,Grupo,MaterialId,PartidaTarifarioId,Posicion,CodigoInterno,CodigoMaterialSAP,CentroSAPFICO,NumeroNecesidad,TipoServicio,Descripcion,Unidad,Cantidad,PrecioOrigen,
            CodigoMonedaOrigen,CodigoTipoIGVOrigen,ZonasTexto,OrigenPrecio,MotivoPrecioManual)
        SELECT CONVERT(INT,item.[key])+1,dato.PartidaId,dato.Grupo,dato.MaterialId,dato.PartidaTarifarioId,dato.Posicion,dato.CodigoInterno,dato.CodigoMaterialSAP,dato.CentroSAPFICO,dato.NumeroNecesidad,dato.TipoServicio,
            NULLIF(LTRIM(RTRIM(dato.Descripcion)),N''),NULLIF(LTRIM(RTRIM(dato.Unidad)),N''),CONVERT(DECIMAL(19,6),dato.Cantidad),TRY_CONVERT(DECIMAL(19,6),dato.PrecioManual),dato.CodigoMonedaOrigen,dato.CodigoTipoIGVOrigen,
            dato.ZonasTexto,dato.OrigenPrecio,NULLIF(LTRIM(RTRIM(dato.MotivoPrecioManual)),N'')
        FROM OPENJSON(@Datos,N'$.Partidas') AS item CROSS APPLY OPENJSON(item.value) WITH
            (PartidaId BIGINT,Grupo NVARCHAR(200),MaterialId BIGINT,PartidaTarifarioId BIGINT,Posicion NVARCHAR(60),CodigoInterno NVARCHAR(60),CodigoMaterialSAP NVARCHAR(60),CentroSAPFICO NVARCHAR(50),NumeroNecesidad NVARCHAR(60),TipoServicio NVARCHAR(150),
            Descripcion NVARCHAR(1500),Unidad NVARCHAR(30),Cantidad NVARCHAR(MAX),PrecioManual NVARCHAR(MAX),CodigoMonedaOrigen NVARCHAR(3),CodigoTipoIGVOrigen NVARCHAR(10),ZonasTexto NVARCHAR(1000),OrigenPrecio NVARCHAR(20),MotivoPrecioManual NVARCHAR(300)) AS dato;
        IF JSON_VALUE(@Datos,N'$.DetalleNeto')=N'true' AND (@TipoIGV<>N'SIN_IGV' OR EXISTS(SELECT 1 FROM #Lineas AS linea WHERE linea.MaterialId IS NULL OR linea.CodigoTipoIGVOrigen<>N'SIN_IGV')) THROW 51317,N'[HYT_VALIDACION] El detalle requiere material y precios sin IGV.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio IS NULL OR linea.OrigenPrecio NOT IN (N'TARIFARIO',N'MANUAL'))
            THROW 51315,N'[HYT_VALIDACION] Selecciona el origen de precio de cada partida.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.PartidaAnterior IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS original WHERE original.PartidaId=linea.PartidaAnterior AND original.RevisionId=@RevisionOrigen))
            THROW 51315,N'[HYT_VALIDACION] Una partida anterior no pertenece a esta revisión.',1;
        IF EXISTS (SELECT linea.PartidaAnterior FROM #Lineas AS linea WHERE linea.PartidaAnterior IS NOT NULL GROUP BY linea.PartidaAnterior HAVING COUNT_BIG(1)>1)
            THROW 51315,N'[HYT_VALIDACION] No se puede repetir el identificador de una partida anterior.',1;
        UPDATE linea SET Conservar=1,MaterialId=original.MaterialId,PrecioOrigen=original.PrecioOrigen,CodigoMonedaOrigen=original.CodigoMonedaOrigen,CodigoTipoIGVOrigen=original.CodigoTipoIGVOrigen,
            Descripcion=COALESCE(linea.Descripcion,original.Descripcion),Unidad=COALESCE(linea.Unidad,original.Unidad),Posicion=COALESCE(linea.Posicion,original.Posicion),CodigoInterno=COALESCE(linea.CodigoInterno,original.CodigoInterno),
            CodigoMaterialSAP=COALESCE(linea.CodigoMaterialSAP,original.CodigoMaterialSAP),ZonasTexto=COALESCE(linea.ZonasTexto,original.ZonasTexto)
        FROM #Lineas AS linea INNER JOIN dbo.PresupuestoPartida AS original ON original.PartidaId=linea.PartidaAnterior AND original.RevisionId=@RevisionOrigen
        WHERE linea.OrigenPrecio=N'TARIFARIO' AND original.OrigenPrecio=N'TARIFARIO' AND linea.PartidaTarifarioId=original.PartidaTarifarioId AND original.CodigoMonedaOrigen=@Moneda AND @ClienteAnterior=@ClienteId;
        UPDATE linea SET MaterialId=material.MaterialId,PrecioOrigen=CASE WHEN @Moneda=N'USD' THEN tarifa.PrecioUSD ELSE tarifa.PrecioPEN END,CodigoMonedaOrigen=@Moneda,CodigoTipoIGVOrigen=tarifa.CodigoTipoIGV,
            Descripcion=COALESCE(linea.Descripcion,material.Descripcion),Unidad=COALESCE(linea.Unidad,unidad.Codigo),Posicion=COALESCE(linea.Posicion,tarifa.Posicion,material.Posicion),
            CodigoInterno=COALESCE(linea.CodigoInterno,material.CodigoInterno),CodigoMaterialSAP=COALESCE(linea.CodigoMaterialSAP,tarifa.CodigoMaterialSAP,material.CodigoMaterialSAP),
            TipoServicio=COALESCE(linea.TipoServicio,servicio.Descripcion),NumeroNecesidad=COALESCE(linea.NumeroNecesidad,tarifa.NumeroNecesidad),
            ZonasTexto=COALESCE(linea.ZonasTexto,(SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=tarifa.PartidaId)),Conservar=1
        FROM #Lineas AS linea INNER JOIN dbo.TarifarioPartida AS tarifa WITH (HOLDLOCK) ON tarifa.PartidaId=linea.PartidaTarifarioId AND tarifa.Activo=1
        INNER JOIN dbo.TarifarioVersion AS version WITH (HOLDLOCK) ON version.VersionId=tarifa.VersionId AND version.Estado=N'VIGENTE'
        INNER JOIN dbo.Tarifario AS cabecera WITH (HOLDLOCK) ON cabecera.TarifarioId=version.TarifarioId AND cabecera.Activo=1 AND (cabecera.ClienteId IS NULL OR cabecera.ClienteId=@ClienteId)
        INNER JOIN dbo.ServicioMaterial AS material WITH (HOLDLOCK) ON material.MaterialId=tarifa.MaterialId AND material.Activo=1
        INNER JOIN dbo.UnidadMedida AS unidad WITH (HOLDLOCK) ON unidad.Id=material.UnidadMedidaId AND unidad.Activo=1
        LEFT JOIN dbo.TipoServicio AS servicio ON servicio.Id=tarifa.TipoServicioId
        WHERE linea.OrigenPrecio=N'TARIFARIO' AND linea.Conservar=0 AND (version.VigenciaDesde IS NULL OR version.VigenciaDesde<=@Emision) AND (version.VigenciaHasta IS NULL OR version.VigenciaHasta>=@Emision);
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio=N'TARIFARIO' AND linea.Conservar=0)
            THROW 51316,N'[HYT_VALIDACION] Una tarifa no está disponible para este cliente o fecha. Revisa la selección.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.Descripcion IS NULL OR linea.Unidad IS NULL OR linea.CodigoMonedaOrigen IS NULL OR linea.CodigoMonedaOrigen<>@Moneda
            OR linea.CodigoTipoIGVOrigen IS NULL OR linea.CodigoTipoIGVOrigen NOT IN (N'CON_IGV',N'SIN_IGV') OR (linea.OrigenPrecio=N'MANUAL' AND linea.PrecioOrigen IS NOT NULL AND linea.MotivoPrecioManual IS NULL))
            THROW 51317,N'[HYT_VALIDACION] Completa descripción, unidad, moneda, tratamiento IGV y motivo de precio manual.',1;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio=N'MANUAL' AND linea.MaterialId IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.ServicioMaterial AS material WHERE material.MaterialId=linea.MaterialId))
            THROW 51317,N'[HYT_VALIDACION] El material indicado no existe.',1;
        UPDATE linea SET CentroSAPFICO=tienda.CentroSAPFICO,Posicion=COALESCE(material.CodigoInterno,linea.CodigoInterno),CodigoInterno=COALESCE(material.CodigoInterno,linea.CodigoInterno)
        FROM #Lineas AS linea CROSS JOIN dbo.Tienda AS tienda
        LEFT JOIN dbo.ServicioMaterial AS material ON material.MaterialId=linea.MaterialId WHERE tienda.TiendaId=@TiendaId;
        IF EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.OrigenPrecio=N'MANUAL' AND linea.PartidaTarifarioId IS NOT NULL AND NOT EXISTS
            (SELECT 1 FROM dbo.TarifarioPartida AS referencia INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=referencia.VersionId
             INNER JOIN dbo.Tarifario AS cabecera ON cabecera.TarifarioId=version.TarifarioId
             WHERE referencia.PartidaId=linea.PartidaTarifarioId AND (linea.MaterialId IS NULL OR referencia.MaterialId=linea.MaterialId) AND (cabecera.ClienteId IS NULL OR cabecera.ClienteId=@ClienteId)))
            THROW 51317,N'[HYT_VALIDACION] La referencia tarifaria no corresponde al material o cliente.',1;
        UPDATE #Lineas SET PrecioUnitario=CONVERT(DECIMAL(19,6),ROUND(
            CASE WHEN @TipoIGV=CodigoTipoIGVOrigen THEN PrecioOrigen WHEN @TipoIGV=N'CON_IGV' THEN PrecioOrigen*(1+@Tasa) ELSE PrecioOrigen/(1+@Tasa) END,2));
        UPDATE #Lineas SET Total=CONVERT(DECIMAL(19,2),ROUND(Cantidad*PrecioUnitario,2));
        UPDATE #Lineas SET BaseImponible=CASE WHEN @TipoIGV=N'CON_IGV' THEN CONVERT(DECIMAL(19,2),ROUND(Total/(1+@Tasa),2)) ELSE Total END;
        UPDATE #Lineas SET Impuesto=CASE WHEN @TipoIGV=N'CON_IGV' THEN Total-BaseImponible ELSE CONVERT(DECIMAL(19,2),ROUND(BaseImponible*@Tasa,2)) END;
        UPDATE #Lineas SET Total=BaseImponible+Impuesto;
        DECLARE @Subtotal DECIMAL(19,2),@Impuesto DECIMAL(19,2),@Total DECIMAL(19,2);
        IF EXISTS (SELECT 1 FROM #Lineas) AND NOT EXISTS (SELECT 1 FROM #Lineas AS linea WHERE linea.Total IS NULL)
            SELECT @Subtotal=SUM(linea.BaseImponible),@Impuesto=SUM(linea.Impuesto),@Total=SUM(linea.Total) FROM #Lineas AS linea;
        DECLARE @ClienteNombre NVARCHAR(200),@ClienteDocumento NVARCHAR(30),@ClienteDireccion NVARCHAR(600),@TiendaNombre NVARCHAR(200),@TiendaDireccion NVARCHAR(600),@TiendaCodigo NVARCHAR(40),@Empresa NVARCHAR(MAX);
        IF @EsNuevo=0 AND @ClienteAnterior=@ClienteId AND @TiendaAnterior=@TiendaId
            SELECT @ClienteNombre=r.ClienteNombre,@ClienteDocumento=r.ClienteDocumento,@ClienteDireccion=r.ClienteDireccion,@TiendaNombre=r.TiendaNombre,@TiendaDireccion=r.TiendaDireccion,@TiendaCodigo=r.TiendaCodigo,@Empresa=r.EmpresaSnapshotJson
            FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionOrigen;
        ELSE
            SELECT @ClienteNombre=CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT(cliente.Nombres,N' ',cliente.Apellidos) ELSE cliente.RazonSocial END,@ClienteDocumento=cliente.NumeroIdentificacion,
                @ClienteDireccion=CONCAT_WS(N' / ',NULLIF(cliente.DireccionFiscal,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=cliente.CodigoUbigeo)),@TiendaNombre=tienda.Nombre,@TiendaDireccion=CONCAT_WS(N' / ',NULLIF(tienda.Direccion,N''),(SELECT CONCAT_WS(N' / ',departamento.Nombre,provincia.Nombre,distrito.Nombre) FROM dbo.UbigeoDistritos AS distrito INNER JOIN dbo.UbigeoProvincias AS provincia ON provincia.CodigoProvincia=distrito.CodigoProvincia INNER JOIN dbo.UbigeoDepartamentos AS departamento ON departamento.CodigoDepartamento=distrito.CodigoDepartamento WHERE distrito.CodigoUbigeo=tienda.CodigoUbigeo)),@TiendaCodigo=tienda.Codigo
            FROM dbo.Cliente AS cliente INNER JOIN dbo.Tienda AS tienda ON tienda.ClienteId=cliente.ClienteId AND tienda.TiendaId=@TiendaId WHERE cliente.ClienteId=@ClienteId;
        IF @Empresa IS NULL OR @NuevaRevision=1
            SELECT @Empresa=(SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        IF @TiendaCodigo IS NULL SELECT @TiendaCodigo=tienda.Codigo FROM dbo.Tienda AS tienda WHERE tienda.TiendaId=@TiendaId;
        IF @EsNuevo=1
        BEGIN
            DECLARE @NumeroInterno NVARCHAR(40);
            EXEC dbo.spu_ObtenerNumeroInterno @TipoRegistro=N'PRESUPUESTO',@NumeroInterno=@NumeroInterno OUTPUT;
            DECLARE @RegistroUtc DATETIME2(3)=SYSUTCDATETIME(),@Periodo CHAR(6),@Secuencia BIGINT,@NumeroPresupuesto NVARCHAR(40);
            SET @Periodo=CONVERT(CHAR(6),CONVERT(DATETIME2,@RegistroUtc AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112);
            SELECT @Secuencia=c.Ultimo FROM dbo.PresupuestoCorrelativoMensual AS c WITH(UPDLOCK,HOLDLOCK) WHERE c.Periodo=@Periodo;
            IF @Secuencia IS NULL BEGIN SET @Secuencia=1;INSERT INTO dbo.PresupuestoCorrelativoMensual(Periodo,Ultimo) VALUES(@Periodo,@Secuencia);END;
            ELSE BEGIN SET @Secuencia=@Secuencia+1;UPDATE dbo.PresupuestoCorrelativoMensual SET Ultimo=@Secuencia WHERE Periodo=@Periodo;END;
            SET @NumeroPresupuesto=CONCAT(@Periodo,CASE WHEN @Secuencia<1000000 THEN RIGHT(CONCAT(N'000000',@Secuencia),6) ELSE CONVERT(NVARCHAR(20),@Secuencia) END);
            INSERT INTO dbo.Presupuesto (NumeroInterno,NumeroPresupuesto,FechaRegistro,ClienteId,TiendaId,TipoAtencion,Descripcion,FechaSolicitud,UsuarioRegistro)
            VALUES (@NumeroInterno,@NumeroPresupuesto,@RegistroUtc,@ClienteId,@TiendaId,@Tipo,@Descripcion,@FechaSolicitud,@UsuarioId);
            SET @PresupuestoId=CONVERT(BIGINT,SCOPE_IDENTITY());
        END
        ELSE UPDATE dbo.Presupuesto SET ClienteId=@ClienteId,TiendaId=@TiendaId,TipoAtencion=@Tipo,Descripcion=@Descripcion,FechaSolicitud=@FechaSolicitud,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        IF @EsNuevo=1 OR @NuevaRevision=1
        BEGIN
            SELECT @NumeroRevision=COALESCE(MAX(r.NumeroRevision),0)+1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId;
            INSERT INTO dbo.PresupuestoRevision (PresupuestoId,NumeroRevision,FechaEmision,ClienteNombre,ClienteDocumento,ClienteDireccion,TiendaNombre,TiendaDireccion,CodigoMoneda,CodigoTipoIGV,FechaDesdeTasaIGV,TasaIGVAplicada,UsuarioRegistro)
            VALUES (@PresupuestoId,@NumeroRevision,@Emision,@ClienteNombre,@ClienteDocumento,@ClienteDireccion,@TiendaNombre,@TiendaDireccion,@Moneda,@TipoIGV,@DesdeTasa,@Tasa,@UsuarioId);
            SET @RevisionId=CONVERT(BIGINT,SCOPE_IDENTITY());
        END;
        UPDATE dbo.PresupuestoRevision SET FechaEmision=@Emision,ClienteNombre=@ClienteNombre,ClienteDocumento=@ClienteDocumento,ClienteDireccion=@ClienteDireccion,TiendaNombre=@TiendaNombre,TiendaDireccion=@TiendaDireccion,TiendaCodigo=@TiendaCodigo,
            NumeroCotizacionCliente=JSON_VALUE(@Datos,N'$.Revision.NumeroCotizacionCliente'),FechaEjecucion=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion'),23),Trimestre=JSON_VALUE(@Datos,N'$.Revision.Trimestre'),
            VigenciaTexto=JSON_VALUE(@Datos,N'$.Revision.VigenciaTexto'),CodigoSAPProveedor=JSON_VALUE(@Datos,N'$.Revision.CodigoSAPProveedor'),ContratoMarco=JSON_VALUE(@Datos,N'$.Revision.ContratoMarco'),Especialidad=JSON_VALUE(@Datos,N'$.Revision.Especialidad'),
            Coordinador=JSON_VALUE(@Datos,N'$.Revision.Coordinador'),Encargado=JSON_VALUE(@Datos,N'$.Revision.Encargado'),Consideraciones=(SELECT dato.value FROM OPENJSON(@Datos,N'$.Revision') AS dato WHERE dato.[key]=N'Consideraciones'),
            Firmante=JSON_VALUE(@Datos,N'$.Revision.Firmante'),CargoFirmante=JSON_VALUE(@Datos,N'$.Revision.CargoFirmante'),ZonaTexto=JSON_VALUE(@Datos,N'$.Revision.ZonaTexto'),EmpresaSnapshotJson=@Empresa,DescripcionTrabajo=@Descripcion,
            CodigoMoneda=@Moneda,SimboloMoneda=(SELECT moneda.Simbolo FROM dbo.Moneda AS moneda WHERE moneda.Codigo=@Moneda),CodigoTipoIGV=@TipoIGV,FechaDesdeTasaIGV=@DesdeTasa,TasaIGVAplicada=@Tasa,Subtotal=@Subtotal,Impuesto=@Impuesto,Total=@Total,
            FechaActualizacion=CASE WHEN @EsNuevo=0 AND @NuevaRevision=0 THEN SYSUTCDATETIME() ELSE NULL END,UsuarioActualizacion=CASE WHEN @EsNuevo=0 AND @NuevaRevision=0 THEN @UsuarioId ELSE NULL END
        WHERE RevisionId=@RevisionId;
        DELETE FROM dbo.PresupuestoPartida WHERE RevisionId=@RevisionId;
        INSERT INTO dbo.PresupuestoPartida (RevisionId,NumeroLinea,Grupo,MaterialId,PartidaTarifarioId,Posicion,CodigoInterno,CodigoMaterialSAP,CentroSAPFICO,NumeroNecesidad,TipoServicio,Descripcion,Unidad,Cantidad,
            PrecioOrigen,CodigoMonedaOrigen,CodigoTipoIGVOrigen,PrecioUnitario,BaseImponible,Impuesto,Total,ZonasTexto,OrigenPrecio,MotivoPrecioManual)
        SELECT @RevisionId,linea.NumeroLinea,linea.Grupo,linea.MaterialId,linea.PartidaTarifarioId,linea.Posicion,linea.CodigoInterno,linea.CodigoMaterialSAP,linea.CentroSAPFICO,linea.NumeroNecesidad,linea.TipoServicio,linea.Descripcion,linea.Unidad,linea.Cantidad,
            linea.PrecioOrigen,linea.CodigoMonedaOrigen,linea.CodigoTipoIGVOrigen,linea.PrecioUnitario,linea.BaseImponible,linea.Impuesto,linea.Total,linea.ZonasTexto,linea.OrigenPrecio,linea.MotivoPrecioManual FROM #Lineas AS linea;
        DECLARE @Posterior NVARCHAR(MAX),@PartidasNuevas NVARCHAR(MAX),@Cambios NVARCHAR(MAX);
        SELECT @Posterior=(SELECT r.RevisionId,r.NumeroRevision,r.FechaEmision,r.NumeroCotizacionCliente,r.ClienteNombre,r.ClienteDocumento,r.ClienteDireccion,r.TiendaNombre,r.TiendaDireccion,r.FechaEjecucion,r.Trimestre,r.VigenciaTexto,
            r.CodigoSAPProveedor,r.ContratoMarco,r.Especialidad,r.Coordinador,r.Encargado,r.Consideraciones,r.Firmante,r.CargoFirmante,r.CodigoMoneda,r.CodigoTipoIGV,r.FechaDesdeTasaIGV,r.TasaIGVAplicada,r.Subtotal,r.Impuesto,r.Total,r.ZonaTexto,r.EmpresaSnapshotJson,r.DescripcionTrabajo,r.SimboloMoneda,r.TiendaCodigo,r.Estado
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionId;
        SELECT @PartidasNuevas=(SELECT p.NumeroLinea,p.Grupo,p.MaterialId,p.PartidaTarifarioId,p.Posicion,p.CodigoInterno,p.CodigoMaterialSAP,p.CentroSAPFICO,p.NumeroNecesidad,p.TipoServicio,p.Descripcion,p.Unidad,p.Cantidad,
            p.PrecioOrigen,p.CodigoMonedaOrigen,p.CodigoTipoIGVOrigen,p.PrecioUnitario,p.BaseImponible,p.Impuesto,p.Total,p.ZonasTexto,p.OrigenPrecio,p.MotivoPrecioManual FROM dbo.PresupuestoPartida AS p WHERE p.RevisionId=@RevisionId ORDER BY p.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES);
        SET @Cambios=(SELECT @RevisionOrigen AS RevisionOrigen,JSON_QUERY(@AnteriorRevision) AS Anterior,JSON_QUERY(@AnteriorPartidas) AS PartidasAnteriores,JSON_QUERY(@Posterior) AS Nuevo,JSON_QUERY(@PartidasNuevas) AS PartidasNuevas FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        SET @Registro=CONVERT(NVARCHAR(100),@RevisionId);
        SET @Accion=CASE WHEN @EsNuevo=1 THEN N'CREAR' WHEN @NuevaRevision=1 THEN N'NUEVA_REVISION' ELSE N'ACTUALIZAR' END;
        DECLARE @EstadoGuardado NVARCHAR(20)=CASE WHEN @EsNuevo=1 OR @NuevaRevision=1 THEN N'BORRADOR' ELSE @Estado END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,@Accion,@Estado,@EstadoGuardado,NULL,NULL,@Cambios;
        SET @Registro=CONVERT(NVARCHAR(100),@PresupuestoId);
        SET @Cambios=(SELECT JSON_QUERY(@Anterior) AS Anterior,@ClienteId AS ClienteId,@TiendaId AS TiendaId,@Tipo AS TipoAtencion,@Descripcion AS Descripcion,@FechaSolicitud AS FechaSolicitud,@RevisionId AS RevisionId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'Presupuesto',@Registro,@Accion,NULL,NULL,NULL,NULL,@Cambios;
        DROP TABLE #Lineas;
        COMMIT TRANSACTION;
        SELECT p.PresupuestoId,r.RevisionId,p.NumeroPresupuesto AS NumeroInterno,r.NumeroRevision,p.VersionFila AS VersionPresupuesto,r.VersionFila AS VersionRevision
        FROM dbo.Presupuesto AS p INNER JOIN dbo.PresupuestoRevision AS r ON r.PresupuestoId=p.PresupuestoId AND r.RevisionId=@RevisionId WHERE p.PresupuestoId=@PresupuestoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        IF OBJECT_ID(N'tempdb..#Lineas') IS NOT NULL DROP TABLE #Lineas;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Presupuesto_Listar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_Listar
    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),
    @Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(20)=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL,@TipoAtencion NVARCHAR(20)=NULL,@PendienteGuia BIT=0,@PendienteOC BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde
            THROW 51301,N'[HYT_VALIDACION] Revisa página y rango de fechas.',1;
        IF @TipoAtencion IS NOT NULL AND @TipoAtencion NOT IN(N'INVITACION',N'EMERGENCIA') THROW 51301,N'[HYT_VALIDACION] Tipo de atención no válido.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT p.PresupuestoId,p.NumeroPresupuesto AS NumeroInterno,p.Descripcion,p.TipoAtencion,p.ClienteId,p.TiendaId,
            revision.RevisionId,revision.NumeroRevision,revision.ClienteNombre,revision.TiendaNombre,revision.Estado,
            revision.CodigoMoneda,revision.Subtotal,revision.Impuesto,revision.Total,revision.FechaAprobacion,
            p.FechaRegistro,creador.Nombres AS UsuarioRegistro,p.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,
            oc.OrdenCompraId,oc.NumeroOCCliente,oc.Estado AS EstadoOC,oc.RevisionId AS RevisionOC,guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,ejecucion.EjecucionId,
            (SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=guia.GuiaId) AS AdjuntosGuia
        INTO #Pagina FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Subtotal,r.Impuesto,r.Total,r.FechaAprobacion,r.FechaActualizacion,r.UsuarioActualizacion
            FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=p.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=p.UsuarioActualizacion
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.Ejecucion AS ejecucion ON ejecucion.PresupuestoId=p.PresupuestoId
        WHERE p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva AND (@TipoAtencion IS NULL OR p.TipoAtencion=@TipoAtencion) AND (@PendienteGuia=0 OR guia.GuiaId IS NULL) AND (@PendienteOC=0 OR oc.OrdenCompraId IS NULL)
          AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR p.NumeroPresupuesto LIKE @Patron ESCAPE N'~' OR p.Descripcion LIKE @Patron ESCAPE N'~' OR revision.ClienteNombre LIKE @Patron ESCAPE N'~' OR revision.TiendaNombre LIKE @Patron ESCAPE N'~');
        SELECT pagina.PresupuestoId,pagina.NumeroInterno,pagina.Descripcion,pagina.TipoAtencion,pagina.ClienteId,pagina.TiendaId,
            pagina.RevisionId,pagina.NumeroRevision,pagina.ClienteNombre,pagina.TiendaNombre,pagina.Estado,pagina.CodigoMoneda,pagina.Subtotal,pagina.Impuesto,pagina.Total,pagina.FechaAprobacion,
            pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.OrdenCompraId,pagina.NumeroOCCliente,pagina.GuiaId,pagina.NumeroActa,pagina.EstadoGuia,pagina.AdjuntosGuia,pagina.EjecucionId
        FROM #Pagina AS pagina WHERE (@Estado IS NULL OR pagina.Estado=@Estado) ORDER BY pagina.FechaRegistro DESC,pagina.PresupuestoId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
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

-- Fuente: StoreProcedure/spu_Presupuesto_ListarImpresiones.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_ListarImpresiones
    @UsuarioId NVARCHAR(128),@RevisionId BIGINT,@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar impresiones.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
        SELECT impresion.ImpresionId,impresion.RevisionId,revision.PresupuestoId,impresion.Formato,impresion.VersionFormato,impresion.FechaRegistro,usuario.Nombres AS Usuario
        FROM dbo.ImpresionPresupuesto AS impresion INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=impresion.RevisionId
        INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=impresion.UsuarioRegistro WHERE impresion.RevisionId=@RevisionId
        ORDER BY impresion.ImpresionId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.ImpresionPresupuesto AS impresion WHERE impresion.RevisionId=@RevisionId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Presupuesto_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_Obtener
    @UsuarioId NVARCHAR(128),@Id BIGINT,@RevisionId BIGINT=NULL,@PaginaAuditoria INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
        IF @RevisionId IS NULL SELECT TOP (1) @RevisionId=r.RevisionId FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@Id ORDER BY r.NumeroRevision DESC;
        IF @RevisionId IS NULL BEGIN SELECT CONVERT(NVARCHAR(MAX),NULL) AS Datos;RETURN;END;
        DECLARE @Referencias TABLE(Entidad NVARCHAR(128) NOT NULL,RegistroId NVARCHAR(100) NOT NULL,PRIMARY KEY(Entidad,RegistroId));
        DECLARE @Permisos TABLE(Codigo NVARCHAR(100) NOT NULL PRIMARY KEY);
        INSERT INTO @Permisos(Codigo) SELECT DISTINCT permiso.CodigoPermiso FROM dbo.AspNetUserRoles AS rol INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE rol.UserId=@UsuarioId;
        INSERT INTO @Referencias(Entidad,RegistroId) VALUES(N'Presupuesto',CONVERT(NVARCHAR(100),@Id)),(N'PresupuestoRevision',CONVERT(NVARCHAR(100),@RevisionId));
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'OrdenCompraCliente',CONVERT(NVARCHAR(100),oc.OrdenCompraId) FROM dbo.OrdenCompraCliente AS oc WHERE oc.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'OC_CLIENTE.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'GuiaConformidad',CONVERT(NVARCHAR(100),guia.GuiaId) FROM dbo.GuiaConformidad AS guia WHERE guia.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'GUIAS.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'Ejecucion',CONVERT(NVARCHAR(100),ejecucion.EjecucionId) FROM dbo.Ejecucion AS ejecucion WHERE ejecucion.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'EJECUCION.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId) SELECT N'ConfirmacionFacturacion',CONVERT(NVARCHAR(100),confirmacion.ConfirmacionId) FROM dbo.ConfirmacionFacturacion AS confirmacion INNER JOIN dbo.OrdenCompraCliente AS oc ON oc.OrdenCompraId=confirmacion.OrdenCompraId WHERE oc.PresupuestoId=@Id AND EXISTS(SELECT 1 FROM @Permisos AS permisoVista WHERE permisoVista.Codigo=N'CONFIRMACIONES.CONSULTAR');
        INSERT INTO @Referencias(Entidad,RegistroId)
        SELECT N'ArchivoAdjunto',CONVERT(NVARCHAR(100),archivo.ArchivoId) FROM dbo.ArchivoAdjunto AS archivo
        WHERE EXISTS(SELECT 1 FROM dbo.OrdenCompraArchivo AS vinculo INNER JOIN @Referencias AS referencia ON referencia.Entidad=N'OrdenCompraCliente' AND referencia.RegistroId=CONVERT(NVARCHAR(100),vinculo.OrdenCompraId) WHERE vinculo.ArchivoId=archivo.ArchivoId)
           OR EXISTS(SELECT 1 FROM dbo.GuiaArchivo AS vinculo INNER JOIN @Referencias AS referencia ON referencia.Entidad=N'GuiaConformidad' AND referencia.RegistroId=CONVERT(NVARCHAR(100),vinculo.GuiaId) WHERE vinculo.ArchivoId=archivo.ArchivoId)
           OR EXISTS(SELECT 1 FROM dbo.ConfirmacionArchivo AS vinculo INNER JOIN @Referencias AS referencia ON referencia.Entidad=N'ConfirmacionFacturacion' AND referencia.RegistroId=CONVERT(NVARCHAR(100),vinculo.ConfirmacionId) WHERE vinculo.ArchivoId=archivo.ArchivoId);
        SELECT (SELECT p.PresupuestoId,p.NumeroPresupuesto AS NumeroInterno,p.ClienteId,p.TiendaId,p.TipoAtencion,p.Descripcion,p.FechaSolicitud,p.FechaRegistro,p.FechaActualizacion,p.VersionFila,
            creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,contactosTienda.Nombres AS ContactosTienda,
            JSON_QUERY((SELECT revision.RevisionId, revision.PresupuestoId, revision.NumeroRevision, revision.FechaEmision, revision.NumeroCotizacionCliente, revision.ClienteNombre, revision.ClienteDocumento, revision.ClienteDireccion, revision.TiendaNombre, revision.TiendaDireccion, revision.FechaEjecucion, revision.Trimestre, revision.VigenciaTexto, revision.CodigoSAPProveedor, revision.ContratoMarco, revision.Especialidad, revision.Coordinador, revision.Encargado, revision.Consideraciones, revision.Firmante, revision.CargoFirmante, revision.CodigoMoneda, revision.CodigoTipoIGV, revision.FechaDesdeTasaIGV, revision.TasaIGVAplicada, revision.Estado, revision.Subtotal, revision.Impuesto, revision.Total, revision.FechaRegistro, revision.UsuarioRegistro, revision.FechaActualizacion, revision.UsuarioActualizacion, revision.VersionFila, revision.ZonaTexto, revision.EmpresaSnapshotJson, revision.PoliticaCalculo, revision.FechaAprobacion, revision.DescripcionTrabajo, revision.SimboloMoneda, revision.TiendaCodigo FROM dbo.PresupuestoRevision AS revision WHERE revision.RevisionId=@RevisionId AND revision.PresupuestoId=p.PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Revision,
            JSON_QUERY((SELECT partida.PartidaId, partida.NumeroLinea, partida.Grupo, partida.MaterialId, partida.PartidaTarifarioId, partida.Posicion, partida.CodigoInterno, partida.CodigoMaterialSAP, COALESCE(NULLIF(LTRIM(RTRIM(partida.CentroSAPFICO)),N''),tienda.CentroSAPFICO) AS CentroSAPFICO, partida.NumeroNecesidad, partida.TipoServicio, partida.Descripcion, partida.Unidad, partida.Cantidad, partida.PrecioOrigen, partida.CodigoMonedaOrigen, partida.CodigoTipoIGVOrigen, partida.PrecioUnitario, partida.BaseImponible, partida.Impuesto, partida.Total, partida.ZonasTexto, partida.OrigenPrecio, partida.MotivoPrecioManual FROM dbo.PresupuestoPartida AS partida WHERE partida.RevisionId=@RevisionId ORDER BY partida.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Partidas,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.Estado,r.FechaEmision,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,oc.RevisionId AS RevisionOC,guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia
        FROM dbo.Presupuesto AS p INNER JOIN dbo.Tienda AS tienda ON tienda.TiendaId=p.TiendaId
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=p.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=p.UsuarioActualizacion
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        OUTER APPLY (
            SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),NULLIF(LTRIM(RTRIM(contacto.Nombre)),N'')),N' / ') WITHIN GROUP(ORDER BY contacto.ContactoId) AS Nombres
            FROM dbo.ContactoCliente AS contacto
            WHERE contacto.ClienteId=p.ClienteId AND contacto.TiendaId=p.TiendaId AND contacto.Activo=1
        ) AS contactosTienda
        WHERE p.PresupuestoId=@Id AND EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionId AND r.PresupuestoId=@Id)
        FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson,evento.Entidad
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE EXISTS(SELECT 1 FROM @Referencias AS referencia WHERE referencia.Entidad=evento.Entidad AND referencia.RegistroId=evento.RegistroId)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE EXISTS(SELECT 1 FROM @Referencias AS referencia WHERE referencia.Entidad=evento.Entidad AND referencia.RegistroId=evento.RegistroId);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Presupuesto_ObtenerImpresion.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_ObtenerImpresion
    @UsuarioId NVARCHAR(128),@ImpresionId BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar impresiones.',1;
        SELECT impresion.ImpresionId,impresion.RevisionId,revision.PresupuestoId,impresion.Formato,impresion.VersionFormato,impresion.FechaRegistro,
            usuario.Nombres AS Usuario,impresion.ContenidoHTML,impresion.SHA256
        FROM dbo.ImpresionPresupuesto AS impresion INNER JOIN dbo.PresupuestoRevision AS revision ON revision.RevisionId=impresion.RevisionId
        INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=impresion.UsuarioRegistro WHERE impresion.ImpresionId=@ImpresionId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_Presupuesto_RegistrarImpresion.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_RegistrarImpresion
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@RevisionId BIGINT,@VersionOriginal BINARY(8),
    @Formato NVARCHAR(10),@VersionFormato NVARCHAR(20),@ContenidoHTML NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para generar impresiones.',1;
        IF @Formato IS NULL OR @Formato NOT IN (N'ANTERIOR',N'NUEVO') OR NULLIF(@VersionFormato,N'') IS NULL OR @ContenidoHTML IS NULL
            OR LEN(@ContenidoHTML)<100 OR LEN(@ContenidoHTML)>2000000 OR LOWER(LEFT(@ContenidoHTML,15))<>N'<!doctype html>'
            THROW 51322,N'[HYT_VALIDACION] Formato o contenido de impresión no válido.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Estado NVARCHAR(20);
        SELECT @Estado=revision.Estado FROM dbo.PresupuestoRevision AS revision WITH (UPDLOCK,HOLDLOCK) WHERE revision.RevisionId=@RevisionId AND revision.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió mientras se generaba la impresión. Vuelve a consultarla.',1;
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@ContenidoHTML)),2);
        INSERT INTO dbo.ImpresionPresupuesto (RevisionId,Formato,VersionFormato,ContenidoHTML,SHA256,UsuarioRegistro)
        VALUES (@RevisionId,@Formato,@VersionFormato,@ContenidoHTML,@Huella,@UsuarioId);
        DECLARE @ImpresionId BIGINT=CONVERT(BIGINT,SCOPE_IDENTITY()),@Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),@Cambios NVARCHAR(MAX);
        SET @Cambios=(SELECT @ImpresionId AS ImpresionId,@Formato AS Formato,@VersionFormato AS VersionFormato,@Huella AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,N'GENERAR_IMPRESION',@Estado,@Estado,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @ImpresionId AS ImpresionId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_ConfiguracionHYT_Guardar.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConfiguracionHYT_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Datos NVARCHAR(MAX),@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIGURACION.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para editar la configuración.',1;
        IF ISJSON(@Datos)<>1 THROW 51303,N'[HYT_VALIDACION] Datos no válidos.',1;
        DECLARE @Ruc NVARCHAR(20)=JSON_VALUE(@Datos,N'$.Ruc'),@Razon NVARCHAR(200)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.RazonSocial'))),N'');
        IF @Ruc IS NULL OR LEN(@Ruc)<>11 OR @Ruc LIKE N'%[^0-9]%' OR @Razon IS NULL
            THROW 51305,N'[HYT_VALIDACION] Completa RUC de once dígitos y razón social.',1;
        IF EXISTS (SELECT 1 FROM OPENJSON(@Datos) AS campo INNER JOIN (VALUES (N'Ruc',11),(N'RazonSocial',200),(N'Direccion',300),(N'CodigoUbigeo',6),(N'Telefono',50),(N'Correo',256),(N'NombreFirmante',150),(N'CargoFirmante',150),(N'ConsideracionesPresupuesto',20000),(N'SubtituloCotizacion',2000)) AS limite(Nombre,Largo)
            ON limite.Nombre=campo.[key] COLLATE DATABASE_DEFAULT WHERE LEN(campo.value)>limite.Largo)
            THROW 51306,N'[HYT_VALIDACION] Un campo supera su longitud permitida.',1;
        DECLARE @CodigoUbigeo CHAR(6)=NULLIF(LTRIM(RTRIM(JSON_VALUE(@Datos,N'$.CodigoUbigeo'))),N'');
        IF @CodigoUbigeo IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.UbigeoDistritos AS distrito WHERE distrito.CodigoUbigeo=@CodigoUbigeo AND distrito.Activo=1) THROW 51305,N'[HYT_VALIDACION] Selecciona un ubigeo válido.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@Existe BIT=0;
        SELECT @Existe=1,@Anterior=(SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.CodigoUbigeo,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.SubtituloCotizacion,empresa.ConsideracionesPresupuesto FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ConfiguracionHYT AS empresa WITH (UPDLOCK,HOLDLOCK) WHERE empresa.Id=1;
        IF @Existe=1 AND NOT EXISTS (SELECT 1 FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1 AND empresa.VersionFila=@VersionOriginal)
            THROW 51309,N'[HYT_CONCURRENCIA] La configuración cambió. Vuelve a consultarla.',1;
        IF @Existe=0 AND @VersionOriginal IS NOT NULL THROW 51309,N'[HYT_CONCURRENCIA] La configuración ya no existe. Vuelve a consultarla.',1;
        IF @Existe=0 INSERT INTO dbo.ConfiguracionHYT (Id,Ruc,RazonSocial,UsuarioRegistro) VALUES (1,@Ruc,@Razon,@UsuarioId);
        UPDATE dbo.ConfiguracionHYT SET SubtituloCotizacion=CASE WHEN EXISTS(SELECT 1 FROM OPENJSON(@Datos) AS campo WHERE campo.[key]=N'SubtituloCotizacion') THEN JSON_VALUE(@Datos,N'$.SubtituloCotizacion') ELSE SubtituloCotizacion END,Ruc=@Ruc,RazonSocial=@Razon,Direccion=JSON_VALUE(@Datos,N'$.Direccion'),CodigoUbigeo=@CodigoUbigeo,Telefono=JSON_VALUE(@Datos,N'$.Telefono'),Correo=JSON_VALUE(@Datos,N'$.Correo'),
            NombreFirmante=JSON_VALUE(@Datos,N'$.NombreFirmante'),CargoFirmante=JSON_VALUE(@Datos,N'$.CargoFirmante'),ConsideracionesPresupuesto=(SELECT dato.value FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'ConsideracionesPresupuesto'),
            FechaActualizacion=CASE WHEN @Existe=1 THEN SYSUTCDATETIME() ELSE NULL END,UsuarioActualizacion=CASE WHEN @Existe=1 THEN @UsuarioId ELSE NULL END WHERE Id=1;
        SELECT @Nuevo=(SELECT empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.CodigoUbigeo,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.SubtituloCotizacion,empresa.ConsideracionesPresupuesto FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1;
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),@Accion NVARCHAR(60)=CASE WHEN @Existe=0 THEN N'CREAR' ELSE N'ACTUALIZAR' END;
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfiguracionHYT',N'1',@Accion,NULL,NULL,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_ConfiguracionHYT_Obtener.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConfiguracionHYT_Obtener
    @UsuarioId NVARCHAR(128),@Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIGURACION.CONSULTAR')
            IF NOT (EXISTS(SELECT 1 FROM dbo.AspNetUsers AS u INNER JOIN dbo.AspNetUserRoles AS r ON r.UserId=u.Id INNER JOIN dbo.RolPermiso AS p ON p.RoleId=r.RoleId WHERE u.Id=@UsuarioId AND u.Activo=1 AND p.CodigoPermiso=N'PRESUPUESTOS.CONSULTAR')
                AND EXISTS(SELECT 1 FROM dbo.AspNetUsers AS u INNER JOIN dbo.AspNetUserRoles AS r ON r.UserId=u.Id INNER JOIN dbo.RolPermiso AS p ON p.RoleId=r.RoleId WHERE u.Id=@UsuarioId AND u.Activo=1 AND p.CodigoPermiso=N'REPORTES.CONSULTAR'))
                THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar la configuración.',1;
        IF @Pagina<1 OR @Pagina>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
        SELECT (SELECT empresa.Id,empresa.Ruc,empresa.RazonSocial,empresa.Direccion,empresa.CodigoUbigeo,empresa.Telefono,empresa.Correo,empresa.NombreFirmante,empresa.CargoFirmante,empresa.SubtituloCotizacion,empresa.ConsideracionesPresupuesto,empresa.FirmaPresupuestoJson,
            empresa.FechaRegistro,creador.Nombres AS UsuarioRegistro,empresa.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,empresa.VersionFila
            FROM dbo.ConfiguracionHYT AS empresa INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=empresa.UsuarioRegistro
            LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=empresa.UsuarioActualizacion WHERE empresa.Id=1
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.Accion,usuario.Nombres AS Usuario,evento.FechaRegistro,evento.EstadoAnterior,evento.EstadoNuevo,evento.CambiosJson
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE evento.Entidad=N'ConfiguracionHYT' AND evento.RegistroId=N'1'
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento WHERE evento.Entidad=N'ConfiguracionHYT' AND evento.RegistroId=N'1';
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/015_PermisosEtapa5.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_AyudaTarifario TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_CambiarEstado TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_DatosIniciales TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_ListarImpresiones TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_Obtener TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_ObtenerImpresion TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_Presupuesto_RegistrarImpresion TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ConfiguracionHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ConfiguracionHYT_Obtener TO hyt_app_executor;
GO

