-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   22/09/2026
-- Description:   Limpia los datos operativos y comerciales de Dbsisadm, conserva maestros y usuarios de plataforma, y reinicia las identidades de las tablas vaciadas.
-- =============================================
-- Firma: FRANCO LARA - 22/09/2026 | Agrega la limpieza integral de SisAdm preservando catalogos, maestros, roles, modulos y los usuarios OwnerPlataforma/SuperAdmin.

/*
    EJECUTAR SOLO EN UNA COPIA DE DESARROLLO O PRUEBAS.

    Se conserva:
    - Catalogos y maestros: ADM_Moneda, ADM_ParametroMaestro, ADM_TipoComprobante,
      ADM_TipoPercepcion, ADM_DetraccionSunat, CON_Bancos, CON_CentroCosto,
      CON_OrigenMaestro, CON_PlanCuentaMaestro, CON_CuentaDestinoReglaMaestro,
      CON_CuentaDestinoReglaDetalleMaestro, CON_ConfiguracionContabilizacionMaestro,
      CON_TipoAfectacionIGV, CON_TipoImpuesto, SEG_ModuloSistema, SEG_RolCuenta,
      SEG_RolCuentaPermiso, AspNetRoles y catalogos SUNAT/ubigeo.
    - Usuarios con el rol OwnerPlataforma o SuperAdmin, incluidos sus perfiles,
      roles y datos de Identity.

    DBCC CHECKIDENT (..., RESEED, 0) deja el siguiente IDENTITY en 1.
*/

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    DECLARE @UsuariosConservados TABLE
    (
        Id NVARCHAR(450) NOT NULL PRIMARY KEY
    );

    INSERT INTO @UsuariosConservados (Id)
    SELECT DISTINCT u.Id
    FROM dbo.AspNetUsers AS u
    INNER JOIN dbo.AspNetUserRoles AS ur
        ON ur.UserId = u.Id
    INNER JOIN dbo.AspNetRoles AS r
        ON r.Id = ur.RoleId
    WHERE UPPER(COALESCE(r.NormalizedName, r.Name)) IN (N'OWNERPLATAFORMA', N'SUPERADMIN');

    IF NOT EXISTS (SELECT 1 FROM @UsuariosConservados)
    BEGIN
        THROW 50001, 'No se encontro un usuario con rol OwnerPlataforma o SuperAdmin. La limpieza fue cancelada.', 1;
    END;

    /* Seguridad dependiente de usuarios, empresas y cuentas. */
    DELETE FROM dbo.SEG_UsuarioCuentaPermiso;
    DELETE FROM dbo.SEG_UsuarioEmpresaPermiso;
    DELETE FROM dbo.SEG_UsuarioCuentaAdministradora;
    DELETE FROM dbo.SEG_UsuarioEmpresa;

    /* Operacion bancaria, compras, ventas y procesos contables. */
    DELETE FROM dbo.BAN_MovimientoBancoDetalle;
    DELETE FROM dbo.COM_CompraDetraccion;
    DELETE FROM dbo.COM_CompraPercepcion;
    DELETE FROM dbo.COM_CompraRetencion;
    DELETE FROM dbo.COM_CompraDetalle;
    DELETE FROM dbo.VEN_VentaDetalle;
    DELETE FROM dbo.CON_AjusteCuentaProcesoDetalle;
    DELETE FROM dbo.CON_AperturaProcesoDetalle;
    DELETE FROM dbo.CON_CierreProcesoDetalle;
    DELETE FROM dbo.CON_DiferenciaCambioProcesoDetalle;
    DELETE FROM dbo.CON_ConfiguracionContabilizacionDetalle;
    DELETE FROM dbo.CON_CuentaDestinoReglaDetalle;
    DELETE FROM dbo.CON_AplicacionNotaCredito;

    UPDATE dbo.BAN_MovimientoBanco
    SET IdMovimientoBancoRelacionado = NULL
    WHERE IdMovimientoBancoRelacionado IS NOT NULL;

    DELETE FROM dbo.BAN_MovimientoBanco;
    DELETE FROM dbo.COM_Compra;
    DELETE FROM dbo.VEN_Venta;
    DELETE FROM dbo.CON_AjusteCuentaProceso;
    DELETE FROM dbo.CON_AperturaProceso;
    DELETE FROM dbo.CON_CierreProceso;
    DELETE FROM dbo.CON_DiferenciaCambioProceso;
    DELETE FROM dbo.CON_AsientoDetalle;
    DELETE FROM dbo.CON_Asiento;

    /* Configuracion y entidades propias de cada empresa. */
    DELETE FROM dbo.CON_LibroElectronicoGeneracion;
    DELETE FROM dbo.CON_PLE_PlanContableControl;
    DELETE FROM dbo.CON_CorrelativoAsiento;
    DELETE FROM dbo.CON_DocumentoConfiguracionEmpresa;
    DELETE FROM dbo.CON_TipoImpuestoConfiguracionEmpresa;
    DELETE FROM dbo.CON_BancosConfiguracionEmpresa;
    DELETE FROM dbo.CON_CentroCostoConfiguracionEmpresa;
    DELETE FROM dbo.CON_PeriodoContableEstado;
    DELETE FROM dbo.CON_CuentaDestinoRegla;
    DELETE FROM dbo.CON_ConfiguracionContabilizacion;
    DELETE FROM dbo.CON_Origen;

    UPDATE dbo.CON_PlanCuenta
    SET IdPlanCuentaPadre = NULL
    WHERE IdPlanCuentaPadre IS NOT NULL;

    DELETE FROM dbo.CON_PlanCuenta;
    DELETE FROM dbo.CON_TipoCambio;
    DELETE FROM dbo.ADM_TipoCambio;
    DELETE FROM dbo.ADM_ParametroEmpresa;
    DELETE FROM dbo.ADM_Cliente;
    DELETE FROM dbo.ADM_Proveedor;
    DELETE FROM dbo.ADM_Persona;

    /* Historial comercial y configuracion de las cuentas administradoras. */
    DELETE FROM dbo.SEG_CuentaAdministradoraSuscripcionPago;
    DELETE FROM dbo.SEG_CuentaAdministradoraSuscripcionMovimiento;
    DELETE FROM dbo.SEG_CuentaAdministradoraSuscripcion;
    DELETE FROM dbo.SEG_CuentaAdministradoraFacturacion;
    DELETE FROM dbo.SEG_CuentaAdministradoraConfiguracion;
    DELETE FROM dbo.SEG_Empresa;
    DELETE FROM dbo.SEG_CuentaAdministradora;

    /* Identity: se conserva exclusivamente la plataforma administrativa. */
    DELETE up
    FROM dbo.SEG_UsuarioPerfil AS up
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM @UsuariosConservados AS uc
        WHERE uc.Id = up.AspNetUserId
    );

    DELETE uc
    FROM dbo.AspNetUserClaims AS uc
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM @UsuariosConservados AS u
        WHERE u.Id = uc.UserId
    );

    DELETE ul
    FROM dbo.AspNetUserLogins AS ul
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM @UsuariosConservados AS u
        WHERE u.Id = ul.UserId
    );

    DELETE ut
    FROM dbo.AspNetUserTokens AS ut
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM @UsuariosConservados AS u
        WHERE u.Id = ut.UserId
    );

    DELETE ur
    FROM dbo.AspNetUserRoles AS ur
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM @UsuariosConservados AS u
        WHERE u.Id = ur.UserId
    );

    DELETE u
    FROM dbo.AspNetUsers AS u
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM @UsuariosConservados AS uc
        WHERE uc.Id = u.Id
    );

    /* Reinicia solo identidades de tablas que quedaron completamente vacias. */
    DECLARE @TablasReseed TABLE
    (
        NombreTabla SYSNAME NOT NULL PRIMARY KEY
    );

    INSERT INTO @TablasReseed (NombreTabla)
    VALUES
        (N'BAN_MovimientoBanco'),
        (N'BAN_MovimientoBancoDetalle'),
        (N'COM_Compra'),
        (N'COM_CompraDetalle'),
        (N'COM_CompraDetraccion'),
        (N'COM_CompraPercepcion'),
        (N'COM_CompraRetencion'),
        (N'VEN_Venta'),
        (N'VEN_VentaDetalle'),
        (N'CON_AjusteCuentaProceso'),
        (N'CON_AjusteCuentaProcesoDetalle'),
        (N'CON_AperturaProceso'),
        (N'CON_AperturaProcesoDetalle'),
        (N'CON_AplicacionNotaCredito'),
        (N'CON_Asiento'),
        (N'CON_AsientoDetalle'),
        (N'CON_BancosConfiguracionEmpresa'),
        (N'CON_CentroCostoConfiguracionEmpresa'),
        (N'CON_CierreProceso'),
        (N'CON_CierreProcesoDetalle'),
        (N'CON_ConfiguracionContabilizacion'),
        (N'CON_ConfiguracionContabilizacionDetalle'),
        (N'CON_CorrelativoAsiento'),
        (N'CON_CuentaDestinoRegla'),
        (N'CON_CuentaDestinoReglaDetalle'),
        (N'CON_DiferenciaCambioProceso'),
        (N'CON_DiferenciaCambioProcesoDetalle'),
        (N'CON_DocumentoConfiguracionEmpresa'),
        (N'CON_LibroElectronicoGeneracion'),
        (N'CON_PeriodoContableEstado'),
        (N'CON_PlanCuenta'),
        (N'CON_PLE_PlanContableControl'),
        (N'CON_TipoCambio'),
        (N'CON_TipoImpuestoConfiguracionEmpresa'),
        (N'ADM_Cliente'),
        (N'ADM_ParametroEmpresa'),
        (N'ADM_Persona'),
        (N'ADM_Proveedor'),
        (N'ADM_TipoCambio'),
        (N'SEG_CuentaAdministradora'),
        (N'SEG_CuentaAdministradoraConfiguracion'),
        (N'SEG_CuentaAdministradoraFacturacion'),
        (N'SEG_CuentaAdministradoraSuscripcion'),
        (N'SEG_CuentaAdministradoraSuscripcionMovimiento'),
        (N'SEG_CuentaAdministradoraSuscripcionPago'),
        (N'SEG_Empresa'),
        (N'SEG_UsuarioCuentaAdministradora'),
        (N'SEG_UsuarioCuentaPermiso'),
        (N'SEG_UsuarioEmpresa'),
        (N'SEG_UsuarioEmpresaPermiso');

    DECLARE @NombreTabla SYSNAME;
    DECLARE @Sql NVARCHAR(MAX);

    DECLARE cursor_reseed CURSOR LOCAL FAST_FORWARD FOR
        SELECT tr.NombreTabla
        FROM @TablasReseed AS tr
        INNER JOIN sys.identity_columns AS ic
            ON ic.object_id = OBJECT_ID(N'dbo.' + tr.NombreTabla);

    OPEN cursor_reseed;
    FETCH NEXT FROM cursor_reseed INTO @NombreTabla;

    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @Sql = N'DBCC CHECKIDENT (N''dbo.' + REPLACE(@NombreTabla, N'''', N'''''') + N''', RESEED, 0);';
        EXEC sys.sp_executesql @Sql;

        FETCH NEXT FROM cursor_reseed INTO @NombreTabla;
    END;

    CLOSE cursor_reseed;
    DEALLOCATE cursor_reseed;

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
    BEGIN
        ROLLBACK TRANSACTION;
    END;

    THROW;
END CATCH;
