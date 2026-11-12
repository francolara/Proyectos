-- Firma: FRANCO LARA | Create date: 10/11/2026 | Asigna el número de presupuesto como número de cotización al guardar.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Guarda borrador o nueva revisión con snapshots, cálculo a dos decimales, concurrencia y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo, consideraciones y detalle neto editable con referencia tarifaria del material validada.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual y nombre completo del tipo de servicio. Edición directa de OBSERVADO conservando su estado y auditoría. Posición desde código interno del material y centro FICO desde tienda.
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Valida por separado el registro y la edición del presupuesto.
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
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @PresupuestoId IS NULL THEN N'PRESUPUESTOS.CREAR' ELSE N'PRESUPUESTOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
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
            NumeroCotizacionCliente=(SELECT p.NumeroPresupuesto FROM dbo.Presupuesto AS p WHERE p.PresupuestoId=@PresupuestoId),FechaEjecucion=TRY_CONVERT(DATE,JSON_VALUE(@Datos,N'$.Revision.FechaEjecucion'),23),Trimestre=JSON_VALUE(@Datos,N'$.Revision.Trimestre'),
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
