"""Procedimientos del mantenimiento de vigencias IGV; solo escribe fuentes."""
import prepare_maintenance as base

permission = base.permission({'permission':'CONFIGURACION'}, 'CONSULTAR')
body = permission + '''        IF @Pagina<1 OR @Pagina>100000000 THROW 51201,N'[HYT_VALIDACION] Página no válida.',1;
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
'''
base.procedure('spu_MantTasaIGV_Listar', 'Lista vigencias de IGV, responsables y auditoría de alta/cierre, veinte por página.',
    '    @UsuarioId NVARCHAR(128), @Busqueda NVARCHAR(200)=NULL, @Pagina INT=1', body)

body = base.permission({'permission':'CONFIGURACION'}, 'EDITAR') + '''        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
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
'''
base.procedure('spu_MantTasaIGV_Agregar', 'Agrega vigencia IGV con concurrencia y auditoría del alta y cierre anterior.',
    '    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @FechaDesde DATE, @Valor DECIMAL(9,6),\n    @Motivo NVARCHAR(1000), @UltimaFechaEsperada DATE=NULL',body,True)
print('Dos procedimientos de vigencias preparados; sin ejecución SQL.')
