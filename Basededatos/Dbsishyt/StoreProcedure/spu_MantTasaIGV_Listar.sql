-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Lista vigencias de IGV, responsables y auditoría de alta/cierre, veinte por página.
-- =============================================
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
