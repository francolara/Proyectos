-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Datos sugeridos de cliente/tienda/proveedor y tasa vigente; sin modificar documentos.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Sugiere Código SAP, direcciones completas con ubigeo y consideraciones de HYT, conservando contratos.
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
