-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/11/2026
-- Description:   Consulta paginada de auditoría y filtro unificado para presupuestos y sus revisiones.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Auditoria_Listar
    @UsuarioId NVARCHAR(128), @Desde DATETIME2(3), @HastaExclusiva DATETIME2(3),
    @Busqueda NVARCHAR(200)=NULL, @ActorFiltro NVARCHAR(128)=NULL,
    @Entidad NVARCHAR(128)=NULL, @Accion NVARCHAR(60)=NULL, @Pagina INT=1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'AUDITORIA.CONSULTAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para consultar auditoría.',1;

        IF @Pagina < 1 OR @Desde IS NULL OR @HastaExclusiva IS NULL OR @HastaExclusiva <= @Desde
            THROW 51700,N'[HYT_VALIDACION] Revisa las fechas o la página solicitada.',1;
        IF @Entidad=N'PresupuestoRevision' SET @Entidad=N'Presupuesto';
        DECLARE @Filtro NVARCHAR(405)=N'%'+REPLACE(REPLACE(REPLACE(REPLACE(TRIM(@Busqueda),N'~',N'~~'),N'%',N'~%'),N'_',N'~_'),N'[',N'~[')+N'%';
        SELECT evento.EventoId, evento.FechaRegistro, evento.Entidad, evento.Accion, evento.RegistroId,
            referencia.Referencia, CONCAT_WS(N' ',actor.Nombres,NULLIF(actor.Apellidos,N'')) AS Usuario
        INTO #Eventos
        FROM dbo.EventoAuditoria AS evento
        INNER JOIN dbo.AspNetUsers AS actor ON actor.Id=evento.UsuarioId
        LEFT JOIN dbo.AspNetUsers AS afectado ON evento.Entidad IN (N'Usuario', N'AspNetUsers', N'Acceso') AND afectado.Id=evento.RegistroId
        LEFT JOIN dbo.AspNetRoles AS perfil ON evento.Entidad IN (N'Perfil', N'AspNetRoles') AND perfil.Id=evento.RegistroId
        LEFT JOIN dbo.PresupuestoRevision AS revision ON evento.Entidad=N'PresupuestoRevision' AND revision.RevisionId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId=CASE
            WHEN evento.Entidad=N'Presupuesto' THEN TRY_CONVERT(BIGINT,evento.RegistroId)
            WHEN evento.Entidad=N'PresupuestoRevision' THEN revision.PresupuestoId END
        LEFT JOIN dbo.OrdenCompraCliente AS oc ON evento.Entidad IN (N'OrdenCompraCliente',N'OC_CLIENTE') AND oc.OrdenCompraId=TRY_CONVERT(BIGINT,evento.RegistroId)
        LEFT JOIN dbo.GuiaConformidad AS guia ON evento.Entidad IN (N'GuiaConformidad',N'GUIA') AND guia.GuiaId=TRY_CONVERT(BIGINT,evento.RegistroId)
        CROSS APPLY (VALUES (COALESCE(afectado.Email,perfil.Name,presupuesto.NumeroPresupuesto,oc.NumeroOCCliente,guia.NumeroActa,evento.RegistroId))) AS referencia(Referencia)
        WHERE evento.FechaRegistro >= @Desde AND evento.FechaRegistro < @HastaExclusiva
            AND (@ActorFiltro IS NULL OR actor.Id=@ActorFiltro OR CHARINDEX(@ActorFiltro,CONCAT_WS(N' ',actor.Nombres,actor.Apellidos,actor.Email))>0)
            AND (@Entidad IS NULL OR evento.Entidad=@Entidad
                OR (@Entidad=N'Presupuesto' AND evento.Entidad=N'PresupuestoRevision'))
            AND (@Accion IS NULL OR evento.Accion=@Accion)
            AND (@Filtro IS NULL OR referencia.Referencia LIKE @Filtro ESCAPE N'~'
                OR evento.RegistroId LIKE @Filtro ESCAPE N'~' OR evento.Motivo LIKE @Filtro ESCAPE N'~'
                OR CONCAT_WS(N' ',actor.Nombres,actor.Apellidos,actor.Email) LIKE @Filtro ESCAPE N'~');

        SELECT (SELECT EventoId, FechaRegistro, Entidad, Accion, RegistroId, Referencia, Usuario
            FROM #Eventos ORDER BY FechaRegistro DESC, EventoId DESC
            OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH) AS Datos;
        SELECT COUNT_BIG(1) AS Total FROM #Eventos;
        DROP TABLE #Eventos;
        SELECT (SELECT modulo.Entidad, evento.Accion
            FROM dbo.EventoAuditoria AS evento
            CROSS APPLY (VALUES (CASE WHEN evento.Entidad=N'PresupuestoRevision'
                THEN N'Presupuesto' ELSE evento.Entidad END)) AS modulo(Entidad)
            WHERE evento.FechaRegistro >= @Desde AND evento.FechaRegistro < @HastaExclusiva
            GROUP BY modulo.Entidad, evento.Accion ORDER BY modulo.Entidad, evento.Accion FOR JSON PATH) AS Opciones;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
