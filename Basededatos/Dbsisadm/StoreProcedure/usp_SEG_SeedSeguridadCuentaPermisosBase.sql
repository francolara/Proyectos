-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Siembra el catalogo de modulos, roles de cuenta y permisos base para seguridad por opcion.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_SeedSeguridadCuentaPermisosBase
    @UsuarioRegistro NVARCHAR(450) = N'sistema'
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Roles TABLE
        (
            CodigoRolCuenta VARCHAR(30) NOT NULL,
            NombreRolCuenta NVARCHAR(100) NOT NULL,
            DescripcionRol NVARCHAR(250) NULL
        );

        INSERT INTO @Roles
        (
            CodigoRolCuenta,
            NombreRolCuenta,
            DescripcionRol
        )
        VALUES
            ('ADMINISTRADORCUENTA', N'Administrador de cuenta', N'Control total de la cuenta administradora, usuarios, empresas y modulos.'),
            ('SUPERVISOR', N'Supervisor', N'Control operativo amplio con capacidad de ver, crear y editar en la mayoria de modulos.'),
            ('OPERADOR', N'Operador', N'Operacion diaria de registros contables y consultas necesarias para la empresa asignada.'),
            ('CONSULTA', N'Consulta', N'Acceso de solo lectura para seguimiento y reportes.');

        UPDATE rc
        SET
            rc.NombreRolCuenta = r.NombreRolCuenta,
            rc.DescripcionRol = r.DescripcionRol,
            rc.Estado = 1
        FROM dbo.SEG_RolCuenta AS rc
        INNER JOIN @Roles AS r
            ON r.CodigoRolCuenta = rc.CodigoRolCuenta;

        INSERT INTO dbo.SEG_RolCuenta
        (
            CodigoRolCuenta,
            NombreRolCuenta,
            DescripcionRol,
            EsRolSistema,
            Estado,
            UsuarioRegistro
        )
        SELECT
            r.CodigoRolCuenta,
            r.NombreRolCuenta,
            r.DescripcionRol,
            1,
            1,
            @UsuarioRegistro
        FROM @Roles AS r
        WHERE NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_RolCuenta AS rc
            WHERE rc.CodigoRolCuenta = r.CodigoRolCuenta
        );

        DECLARE @Modulos TABLE
        (
            CodigoModulo VARCHAR(50) NOT NULL,
            NombreModulo NVARCHAR(150) NOT NULL,
            DescripcionModulo NVARCHAR(250) NULL,
            AlcanceModulo VARCHAR(20) NOT NULL,
            GrupoMenu NVARCHAR(100) NULL,
            OrdenMenu INT NOT NULL
        );

        INSERT INTO @Modulos
        (
            CodigoModulo,
            NombreModulo,
            DescripcionModulo,
            AlcanceModulo,
            GrupoMenu,
            OrdenMenu
        )
        VALUES
            ('DASHBOARD', N'Dashboard', N'Resumen principal del sistema.', 'CUENTA', N'General', 10),
            ('EMPRESAS', N'Empresas', N'Administracion de empresas vinculadas a la cuenta.', 'CUENTA', N'General', 20),
            ('USUARIOS', N'Usuarios', N'Administracion de usuarios, empresas asignadas y permisos.', 'CUENTA', N'General', 30),
            ('CONFIGURACION', N'Configuracion', N'Datos de la cuenta administradora y facturacion.', 'CUENTA', N'General', 40),
            ('MISUSCRIPCION', N'Mi suscripcion', N'Consulta de plan, estado y limites comerciales.', 'CUENTA', N'General', 50),
            ('AYUDA', N'Ayuda', N'Centro de ayuda y soporte.', 'CUENTA', N'General', 60),
            ('PLANCUENTA', N'Plan de cuentas', N'Mantenimiento del plan contable.', 'EMPRESA', N'Mantenimiento', 110),
            ('CENTROCOSTO', N'Centros de costo', N'Mantenimiento de centros de costo.', 'EMPRESA', N'Mantenimiento', 120),
            ('CUENTACORRIENTE', N'Cuentas corrientes', N'Mantenimiento de cuentas corrientes.', 'EMPRESA', N'Mantenimiento', 130),
            ('PERSONAS', N'Personas', N'Mantenimiento de personas, clientes y proveedores.', 'EMPRESA', N'Mantenimiento', 140),
            ('TIPOCAMBIO', N'Tipos de cambio', N'Consulta y mantenimiento de tipo de cambio.', 'EMPRESA', N'Mantenimiento', 150),
            ('ORIGENES', N'Origenes', N'Mantenimiento de origenes contables.', 'EMPRESA', N'Mantenimiento', 160),
            ('CUENTASDESTINO', N'Cuentas destino', N'Reglas de cuentas destino.', 'EMPRESA', N'Mantenimiento', 170),
            ('CONFIGCONTABLE', N'Configuracion contable', N'Configuraciones base de contabilizacion.', 'EMPRESA', N'Mantenimiento', 180),
            ('ASIENTOS', N'Asientos', N'Registro y gestion de asientos contables.', 'EMPRESA', N'Registro', 210),
            ('COMPRAS', N'Compras', N'Registro y provision de compras.', 'EMPRESA', N'Registro', 220),
            ('VENTAS', N'Ventas', N'Registro y provision de ventas.', 'EMPRESA', N'Registro', 230),
            ('CAJABANCOS', N'Caja y Bancos', N'Movimientos de caja y bancos.', 'EMPRESA', N'Registro', 240),
            ('TRANSFERENCIAS', N'Transferencias', N'Transferencias entre cuentas.', 'EMPRESA', N'Registro', 250),
            ('APLICACIONES', N'Aplicaciones', N'Aplicacion de notas de credito y documentos relacionados.', 'EMPRESA', N'Registro', 260),
            ('PROCESOS', N'Procesos', N'Procesos de cierre, apertura, ajuste y diferencia de cambio.', 'EMPRESA', N'Proceso', 310),
            ('REPORTES', N'Reportes', N'Reportes de analisis y libros contables.', 'EMPRESA', N'Reportes', 410),
            ('LIBROELECTRONICO', N'Libros Electronicos', N'Generacion e historial de libros electronicos.', 'EMPRESA', N'Reportes', 420);

        UPDATE ms
        SET
            ms.NombreModulo = m.NombreModulo,
            ms.DescripcionModulo = m.DescripcionModulo,
            ms.AlcanceModulo = m.AlcanceModulo,
            ms.GrupoMenu = m.GrupoMenu,
            ms.OrdenMenu = m.OrdenMenu,
            ms.EsVisibleMenu = 1,
            ms.Estado = 1
        FROM dbo.SEG_ModuloSistema AS ms
        INNER JOIN @Modulos AS m
            ON m.CodigoModulo = ms.CodigoModulo;

        INSERT INTO dbo.SEG_ModuloSistema
        (
            CodigoModulo,
            NombreModulo,
            DescripcionModulo,
            AlcanceModulo,
            GrupoMenu,
            OrdenMenu,
            EsVisibleMenu,
            Estado,
            UsuarioRegistro
        )
        SELECT
            m.CodigoModulo,
            m.NombreModulo,
            m.DescripcionModulo,
            m.AlcanceModulo,
            m.GrupoMenu,
            m.OrdenMenu,
            1,
            1,
            @UsuarioRegistro
        FROM @Modulos AS m
        WHERE NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_ModuloSistema AS ms
            WHERE ms.CodigoModulo = m.CodigoModulo
        );

        DECLARE @Permisos TABLE
        (
            CodigoRolCuenta VARCHAR(30) NOT NULL,
            CodigoModulo VARCHAR(50) NOT NULL,
            PuedeVer BIT NOT NULL,
            PuedeCrear BIT NOT NULL,
            PuedeEditar BIT NOT NULL,
            PuedeEliminar BIT NOT NULL
        );

        INSERT INTO @Permisos
        (
            CodigoRolCuenta,
            CodigoModulo,
            PuedeVer,
            PuedeCrear,
            PuedeEditar,
            PuedeEliminar
        )
        SELECT
            r.CodigoRolCuenta,
            m.CodigoModulo,
            1,
            1,
            1,
            1
        FROM @Roles AS r
        CROSS JOIN @Modulos AS m
        WHERE r.CodigoRolCuenta = 'ADMINISTRADORCUENTA';

        INSERT INTO @Permisos
        (
            CodigoRolCuenta,
            CodigoModulo,
            PuedeVer,
            PuedeCrear,
            PuedeEditar,
            PuedeEliminar
        )
        VALUES
            ('SUPERVISOR', 'DASHBOARD', 1, 0, 0, 0),
            ('SUPERVISOR', 'EMPRESAS', 1, 0, 0, 0),
            ('SUPERVISOR', 'USUARIOS', 1, 1, 1, 0),
            ('SUPERVISOR', 'CONFIGURACION', 1, 0, 1, 0),
            ('SUPERVISOR', 'MISUSCRIPCION', 1, 0, 0, 0),
            ('SUPERVISOR', 'AYUDA', 1, 0, 0, 0),
            ('SUPERVISOR', 'PLANCUENTA', 1, 1, 1, 0),
            ('SUPERVISOR', 'CENTROCOSTO', 1, 1, 1, 0),
            ('SUPERVISOR', 'CUENTACORRIENTE', 1, 1, 1, 0),
            ('SUPERVISOR', 'PERSONAS', 1, 1, 1, 0),
            ('SUPERVISOR', 'TIPOCAMBIO', 1, 1, 1, 0),
            ('SUPERVISOR', 'ORIGENES', 1, 1, 1, 0),
            ('SUPERVISOR', 'CUENTASDESTINO', 1, 1, 1, 0),
            ('SUPERVISOR', 'CONFIGCONTABLE', 1, 1, 1, 0),
            ('SUPERVISOR', 'ASIENTOS', 1, 1, 1, 0),
            ('SUPERVISOR', 'COMPRAS', 1, 1, 1, 0),
            ('SUPERVISOR', 'VENTAS', 1, 1, 1, 0),
            ('SUPERVISOR', 'CAJABANCOS', 1, 1, 1, 0),
            ('SUPERVISOR', 'TRANSFERENCIAS', 1, 1, 1, 0),
            ('SUPERVISOR', 'APLICACIONES', 1, 1, 1, 0),
            ('SUPERVISOR', 'PROCESOS', 1, 1, 1, 0),
            ('SUPERVISOR', 'REPORTES', 1, 0, 0, 0),
            ('SUPERVISOR', 'LIBROELECTRONICO', 1, 1, 1, 0),
            ('OPERADOR', 'DASHBOARD', 1, 0, 0, 0),
            ('OPERADOR', 'AYUDA', 1, 0, 0, 0),
            ('OPERADOR', 'PERSONAS', 1, 1, 1, 0),
            ('OPERADOR', 'TIPOCAMBIO', 1, 1, 1, 0),
            ('OPERADOR', 'ASIENTOS', 1, 1, 1, 0),
            ('OPERADOR', 'COMPRAS', 1, 1, 1, 0),
            ('OPERADOR', 'VENTAS', 1, 1, 1, 0),
            ('OPERADOR', 'CAJABANCOS', 1, 1, 1, 0),
            ('OPERADOR', 'TRANSFERENCIAS', 1, 1, 1, 0),
            ('OPERADOR', 'APLICACIONES', 1, 1, 1, 0),
            ('OPERADOR', 'REPORTES', 1, 0, 0, 0),
            ('CONSULTA', 'DASHBOARD', 1, 0, 0, 0),
            ('CONSULTA', 'AYUDA', 1, 0, 0, 0),
            ('CONSULTA', 'PLANCUENTA', 1, 0, 0, 0),
            ('CONSULTA', 'CENTROCOSTO', 1, 0, 0, 0),
            ('CONSULTA', 'CUENTACORRIENTE', 1, 0, 0, 0),
            ('CONSULTA', 'PERSONAS', 1, 0, 0, 0),
            ('CONSULTA', 'TIPOCAMBIO', 1, 0, 0, 0),
            ('CONSULTA', 'ORIGENES', 1, 0, 0, 0),
            ('CONSULTA', 'CUENTASDESTINO', 1, 0, 0, 0),
            ('CONSULTA', 'CONFIGCONTABLE', 1, 0, 0, 0),
            ('CONSULTA', 'ASIENTOS', 1, 0, 0, 0),
            ('CONSULTA', 'COMPRAS', 1, 0, 0, 0),
            ('CONSULTA', 'VENTAS', 1, 0, 0, 0),
            ('CONSULTA', 'CAJABANCOS', 1, 0, 0, 0),
            ('CONSULTA', 'TRANSFERENCIAS', 1, 0, 0, 0),
            ('CONSULTA', 'APLICACIONES', 1, 0, 0, 0),
            ('CONSULTA', 'PROCESOS', 1, 0, 0, 0),
            ('CONSULTA', 'REPORTES', 1, 0, 0, 0),
            ('CONSULTA', 'LIBROELECTRONICO', 1, 0, 0, 0);

        UPDATE rcp
        SET
            rcp.PuedeVer = p.PuedeVer,
            rcp.PuedeCrear = p.PuedeCrear,
            rcp.PuedeEditar = p.PuedeEditar,
            rcp.PuedeEliminar = p.PuedeEliminar,
            rcp.Estado = 1
        FROM dbo.SEG_RolCuentaPermiso AS rcp
        INNER JOIN dbo.SEG_RolCuenta AS rc
            ON rc.IdRolCuenta = rcp.IdRolCuenta
        INNER JOIN dbo.SEG_ModuloSistema AS ms
            ON ms.IdModuloSistema = rcp.IdModuloSistema
        INNER JOIN @Permisos AS p
            ON p.CodigoRolCuenta = rc.CodigoRolCuenta
           AND p.CodigoModulo = ms.CodigoModulo;

        INSERT INTO dbo.SEG_RolCuentaPermiso
        (
            IdRolCuenta,
            IdModuloSistema,
            PuedeVer,
            PuedeCrear,
            PuedeEditar,
            PuedeEliminar,
            Estado,
            UsuarioRegistro
        )
        SELECT
            rc.IdRolCuenta,
            ms.IdModuloSistema,
            p.PuedeVer,
            p.PuedeCrear,
            p.PuedeEditar,
            p.PuedeEliminar,
            1,
            @UsuarioRegistro
        FROM @Permisos AS p
        INNER JOIN dbo.SEG_RolCuenta AS rc
            ON rc.CodigoRolCuenta = p.CodigoRolCuenta
        INNER JOIN dbo.SEG_ModuloSistema AS ms
            ON ms.CodigoModulo = p.CodigoModulo
        WHERE NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_RolCuentaPermiso AS rcp
            WHERE rcp.IdRolCuenta = rc.IdRolCuenta
              AND rcp.IdModuloSistema = ms.IdModuloSistema
        );

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
