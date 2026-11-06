"""Fuentes incrementales de presupuestos/documentos. No abre conexiones."""
from pathlib import Path
import json

ROOT = Path(__file__).resolve().parents[3]
BASE = ROOT / 'Basededatos/Dbsishyt'
def save(path, text):
    path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))

def permission(code):
    return f'''        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'{code}')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
'''

scripts = ['Tablas/PresupuestoCorrelativoMensual.sql','Script/026_NumeroPresupuestoMensual.sql','Script/027_NombreTipoServicioPresupuesto.sql','Script/014_PresupuestosDocumentos.sql']
def procedure(name, description, parameters, body, transaction=False):
    text=f'''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   {description}
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.{name}
{parameters}
AS
BEGIN
    SET NOCOUNT ON;
'''
    if transaction: text+='    SET XACT_ABORT ON;\n    DECLARE @Propia BIT=0;\n'
    text+='    BEGIN TRY\n'+body+'    END TRY\n    BEGIN CATCH\n'
    if transaction:text+='        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;\n'
    if '#Pagina' in body:text+="        IF OBJECT_ID(N'tempdb..#Pagina') IS NOT NULL DROP TABLE #Pagina;\n"
    text+='''        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
'''
    if name=="spu_Presupuesto_Listar":
        text=text.replace("USE [SistemaWebHYT];","-- Firma: FRANCO LARA | Create date: 05/11/2026 | Filtros cliente y tienda conservados desde dashboard.\nUSE [SistemaWebHYT];",1)
    relative='StoreProcedure/'+name+'.sql'; save(BASE/relative,text); scripts.append(relative)

body=permission('PRESUPUESTOS.CONSULTAR')+'''        IF @Pagina<1 OR @Pagina>100000000 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva<=@Desde
            THROW 51301,N'[HYT_VALIDACION] Revisa página y rango de fechas.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT p.PresupuestoId,p.NumeroInterno,p.Descripcion,p.TipoAtencion,p.ClienteId,p.TiendaId,
            revision.RevisionId,revision.NumeroRevision,revision.ClienteNombre,revision.TiendaNombre,revision.Estado,
            revision.CodigoMoneda,revision.Subtotal,revision.Impuesto,revision.Total,revision.FechaAprobacion,
            p.FechaRegistro,creador.Nombres AS UsuarioRegistro,p.FechaActualizacion,editor.Nombres AS UsuarioActualizacion,
            oc.OrdenCompraId,oc.NumeroOCCliente,oc.Estado AS EstadoOC,oc.RevisionId AS RevisionOC,guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia,
            (SELECT COUNT_BIG(1) FROM dbo.GuiaArchivo AS vinculo INNER JOIN dbo.ArchivoAdjunto AS archivo ON archivo.ArchivoId=vinculo.ArchivoId AND archivo.Retirado=0 WHERE vinculo.GuiaId=guia.GuiaId) AS AdjuntosGuia
        INTO #Pagina FROM dbo.Presupuesto AS p
        OUTER APPLY (SELECT TOP (1) r.RevisionId,r.NumeroRevision,r.ClienteNombre,r.TiendaNombre,r.Estado,r.CodigoMoneda,r.Subtotal,r.Impuesto,r.Total,r.FechaAprobacion,r.FechaActualizacion,r.UsuarioActualizacion
            FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC) AS revision
        INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=p.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=p.UsuarioActualizacion
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        WHERE p.FechaRegistro>=@Desde AND p.FechaRegistro<@HastaExclusiva AND (@Estado IS NULL OR revision.Estado=@Estado)
          AND (@Cliente IS NULL OR revision.ClienteNombre LIKE N'%'+@Cliente+N'%') AND (@Tienda IS NULL OR revision.TiendaNombre LIKE N'%'+@Tienda+N'%')
            AND (@Busqueda IS NULL OR p.NumeroInterno LIKE @Patron ESCAPE N'~' OR p.Descripcion LIKE @Patron ESCAPE N'~' OR revision.ClienteNombre LIKE @Patron ESCAPE N'~' OR revision.TiendaNombre LIKE @Patron ESCAPE N'~');
        SELECT pagina.PresupuestoId,pagina.NumeroInterno,pagina.Descripcion,pagina.TipoAtencion,pagina.ClienteId,pagina.TiendaId,
            pagina.RevisionId,pagina.NumeroRevision,pagina.ClienteNombre,pagina.TiendaNombre,pagina.Estado,pagina.CodigoMoneda,pagina.Subtotal,pagina.Impuesto,pagina.Total,pagina.FechaAprobacion,
            pagina.FechaRegistro,pagina.UsuarioRegistro,pagina.FechaActualizacion,pagina.UsuarioActualizacion,pagina.OrdenCompraId,pagina.NumeroOCCliente,pagina.GuiaId,pagina.NumeroActa,pagina.EstadoGuia,pagina.AdjuntosGuia
        FROM #Pagina AS pagina ORDER BY pagina.FechaRegistro DESC,pagina.PresupuestoId DESC OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
        DROP TABLE #Pagina;
'''
procedure('spu_Presupuesto_Listar','Lista presupuestos por búsqueda y rango UTC; última revisión, OC y guía, veinte por página.',
    '    @UsuarioId NVARCHAR(128),@Desde DATETIME2(3),@HastaExclusiva DATETIME2(3),\n    @Pagina INT=1,@Busqueda NVARCHAR(200)=NULL,@Estado NVARCHAR(20)=NULL,@Cliente NVARCHAR(200)=NULL,@Tienda NVARCHAR(200)=NULL',body)

revision_columns='RevisionId PresupuestoId NumeroRevision FechaEmision NumeroCotizacionCliente ClienteNombre ClienteDocumento ClienteDireccion TiendaNombre TiendaDireccion FechaEjecucion Trimestre VigenciaTexto CodigoSAPProveedor ContratoMarco Especialidad Coordinador Encargado Consideraciones Firmante CargoFirmante CodigoMoneda CodigoTipoIGV FechaDesdeTasaIGV TasaIGVAplicada Estado Subtotal Impuesto Total FechaRegistro UsuarioRegistro FechaActualizacion UsuarioActualizacion VersionFila ZonaTexto EmpresaSnapshotJson PoliticaCalculo FechaAprobacion DescripcionTrabajo SimboloMoneda TiendaCodigo'.split()
line_columns='PartidaId NumeroLinea Grupo MaterialId PartidaTarifarioId Posicion CodigoInterno CodigoMaterialSAP CentroSAPFICO NumeroNecesidad TipoServicio Descripcion Unidad Cantidad PrecioOrigen CodigoMonedaOrigen CodigoTipoIGVOrigen PrecioUnitario BaseImponible Impuesto Total ZonasTexto OrigenPrecio MotivoPrecioManual'.split()
body=permission('PRESUPUESTOS.CONSULTAR')+'''        IF @PaginaAuditoria<1 OR @PaginaAuditoria>100000000 THROW 51301,N'[HYT_VALIDACION] Página no válida.',1;
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
        SELECT (SELECT p.PresupuestoId,p.NumeroInterno,p.ClienteId,p.TiendaId,p.TipoAtencion,p.Descripcion,p.FechaSolicitud,p.FechaRegistro,p.FechaActualizacion,p.VersionFila,
            creador.Nombres AS UsuarioRegistro,editor.Nombres AS UsuarioActualizacion,
            JSON_QUERY((SELECT '''+', '.join('revision.'+c for c in revision_columns)+''' FROM dbo.PresupuestoRevision AS revision WHERE revision.RevisionId=@RevisionId AND revision.PresupuestoId=p.PresupuestoId FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)) AS Revision,
            JSON_QUERY((SELECT '''+', '.join('partida.'+c for c in line_columns)+''' FROM dbo.PresupuestoPartida AS partida WHERE partida.RevisionId=@RevisionId ORDER BY partida.NumeroLinea FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Partidas,
            JSON_QUERY((SELECT r.RevisionId,r.NumeroRevision,r.Estado,r.FechaEmision,r.CodigoMoneda,r.Total FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=p.PresupuestoId ORDER BY r.NumeroRevision DESC FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Revisiones,
            oc.OrdenCompraId,oc.NumeroRegistroInterno AS NumeroInternoOC,oc.NumeroOCCliente,oc.Estado AS EstadoOC,oc.RevisionId AS RevisionOC,guia.GuiaId,guia.NumeroActa,guia.Estado AS EstadoGuia
        FROM dbo.Presupuesto AS p INNER JOIN dbo.AspNetUsers AS creador ON creador.Id=p.UsuarioRegistro
        LEFT JOIN dbo.AspNetUsers AS editor ON editor.Id=p.UsuarioActualizacion
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON oc.PresupuestoId=p.PresupuestoId
        LEFT JOIN dbo.GuiaConformidad AS guia ON guia.PresupuestoId=p.PresupuestoId
        WHERE p.PresupuestoId=@Id AND EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.RevisionId=@RevisionId AND r.PresupuestoId=@Id)
        FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) AS Datos;
        SELECT evento.EventoId,evento.Accion,evento.EstadoAnterior,evento.EstadoNuevo,evento.FechaRegistro,evento.FechaEvento,usuario.Nombres AS Usuario,evento.Motivo,evento.CambiosJson,evento.Entidad
        FROM dbo.EventoAuditoria AS evento INNER JOIN dbo.AspNetUsers AS usuario ON usuario.Id=evento.UsuarioId
        WHERE EXISTS(SELECT 1 FROM @Referencias AS referencia WHERE referencia.Entidad=evento.Entidad AND referencia.RegistroId=evento.RegistroId)
        ORDER BY evento.EventoId DESC OFFSET (CONVERT(BIGINT,@PaginaAuditoria)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM dbo.EventoAuditoria AS evento
        WHERE EXISTS(SELECT 1 FROM @Referencias AS referencia WHERE referencia.Entidad=evento.Entidad AND referencia.RegistroId=evento.RegistroId);
'''
procedure('spu_Presupuesto_Obtener','Recupera revisión y partidas históricas, revisiones y auditoría del expediente paginada y filtrada por permisos.',
    '    @UsuarioId NVARCHAR(128),@Id BIGINT,@RevisionId BIGINT=NULL,@PaginaAuditoria INT=1',body)

body=permission('PRESUPUESTOS.CONSULTAR')+'''        IF @Pagina<1 OR @Pagina>100000000 OR @Fecha IS NULL OR @CodigoMoneda NOT IN (N'PEN',N'USD') THROW 51301,N'[HYT_VALIDACION] Revisa página, fecha y moneda.',1;
        SET @Busqueda=NULLIF(LTRIM(RTRIM(@Busqueda)),N'');
        DECLARE @Patron NVARCHAR(410)=N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[') + N'%';
        SELECT partida.PartidaId AS PartidaTarifarioId,material.MaterialId,material.CodigoInterno,COALESCE(partida.Posicion,material.Posicion) AS Posicion,
            COALESCE(partida.CodigoMaterialSAP,material.CodigoMaterialSAP) AS CodigoMaterialSAP,material.Descripcion,unidad.Codigo AS Unidad,
            tarifa.Codigo AS CodigoTarifario,version.NumeroVersion,partida.CodigoTipoIGV,@CodigoMoneda AS CodigoMoneda,
            CONVERT(DECIMAL(19,2),CASE WHEN @CodigoMoneda=N'USD' THEN partida.PrecioUSD ELSE partida.PrecioPEN END) AS Precio,
            servicio.Codigo AS TipoServicio,partida.NumeroNecesidad,
            (SELECT STRING_AGG(CONVERT(NVARCHAR(MAX),zona.Descripcion),N', ') FROM dbo.TarifarioPartidaZona AS vinculo INNER JOIN dbo.Zona AS zona ON zona.Id=vinculo.ZonaId WHERE vinculo.PartidaId=partida.PartidaId) AS Zonas
        INTO #Pagina FROM dbo.TarifarioPartida AS partida
        INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=partida.VersionId
        INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
        INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=partida.MaterialId
        INNER JOIN dbo.UnidadMedida AS unidad ON unidad.Id=material.UnidadMedidaId
        LEFT JOIN dbo.TipoServicio AS servicio ON servicio.Id=partida.TipoServicioId
        WHERE partida.Activo=1 AND tarifa.Activo=1 AND material.Activo=1 AND unidad.Activo=1 AND version.Estado=N'VIGENTE'
            AND (tarifa.ClienteId IS NULL OR tarifa.ClienteId=@ClienteId)
            AND (version.VigenciaDesde IS NULL OR version.VigenciaDesde<=@Fecha) AND (version.VigenciaHasta IS NULL OR version.VigenciaHasta>=@Fecha)
            AND (@Busqueda IS NULL OR material.CodigoInterno LIKE @Patron ESCAPE N'~' OR material.Descripcion LIKE @Patron ESCAPE N'~' OR partida.CodigoMaterialSAP LIKE @Patron ESCAPE N'~' OR partida.Posicion LIKE @Patron ESCAPE N'~');
        SELECT pagina.PartidaTarifarioId,pagina.MaterialId,pagina.CodigoInterno,pagina.Posicion,pagina.CodigoMaterialSAP,pagina.Descripcion,pagina.Unidad,pagina.CodigoTarifario,pagina.NumeroVersion,
            pagina.CodigoTipoIGV,pagina.CodigoMoneda,pagina.Precio,pagina.TipoServicio,pagina.NumeroNecesidad,pagina.Zonas
        FROM #Pagina AS pagina ORDER BY pagina.CodigoInterno,pagina.PartidaTarifarioId OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS Total FROM #Pagina AS pagina;
        DROP TABLE #Pagina;
'''
procedure('spu_Presupuesto_AyudaTarifario','Ayuda de tarifas publicadas por cliente/fecha y moneda; zonas solo informativas.',
    '    @UsuarioId NVARCHAR(128),@ClienteId BIGINT,@Fecha DATE,@CodigoMoneda NVARCHAR(3)=N\'PEN\',@Busqueda NVARCHAR(200)=NULL,@Pagina INT=1',body)

scripts=['Script/014_PresupuestosDocumentos.sql'] + sorted('StoreProcedure/'+path.name for path in (BASE/'StoreProcedure').glob('spu_Presupuesto_*.sql'))
scripts+=sorted('StoreProcedure/'+path.name for path in (BASE/'StoreProcedure').glob('spu_ConfiguracionHYT_*.sql'))
grant='Script/015_PermisosEtapa5.sql'
granttext='''-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Permisos individuales de presupuestos, etapa 5.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
'''
granttext+=''.join(f'GRANT EXECUTE ON OBJECT::dbo.{Path(path).stem} TO hyt_app_executor;\n' for path in scripts if path.startswith('StoreProcedure/'))
save(BASE/grant,granttext+'GO\n'); scripts.append(grant)
save(BASE/'manifest-etapa5.json',json.dumps({'database':'SistemaWebHYT','requires':'etapa4','scripts':scripts},indent=2)+'\n')
manifest=json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))
prior=[]
for stage in range(1,5):prior+=json.loads((BASE/f'manifest-etapa{stage}.json').read_text(encoding='utf-8'))['scripts']
manifest['scripts']=prior+scripts
for stage in range(6,10):
    if (BASE/f'manifest-etapa{stage}.json').exists():manifest['scripts']+=json.loads((BASE/f'manifest-etapa{stage}.json').read_text(encoding='utf-8'))['scripts']
save(BASE/'manifest.json',json.dumps(manifest,indent=2)+'\n')
print(f'Etapa 5 en desarrollo: {len(scripts)} fuentes registradas; sin ejecución SQL desde el generador.')
