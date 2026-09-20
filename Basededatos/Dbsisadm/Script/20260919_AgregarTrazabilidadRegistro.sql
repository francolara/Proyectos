-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/09/2026
-- Description:   Agrega campos de trazabilidad sin alterar los registros historicos; las altas y actualizaciones se administran desde los procedimientos almacenados.
-- =============================================

SET NOCOUNT ON;

DECLARE @Tablas TABLE (Entidad VARCHAR(30), Tabla SYSNAME, ColumnaId SYSNAME);
INSERT INTO @Tablas (Entidad, Tabla, ColumnaId) VALUES
    ('PLANCUENTA', 'CON_PlanCuenta', 'IdPlanCuenta'),
    ('CENTROCOSTO', 'CON_CentroCostoConfiguracionEmpresa', 'IdCentroCostoConfiguracionEmpresa'),
    ('CUENTACORRIENTE', 'CON_BancosConfiguracionEmpresa', 'IdBancoConfiguracionEmpresa'),
    ('PERSONA', 'ADM_Persona', 'IdPersona'),
    ('TIPOCAMBIO', 'CON_TipoCambio', 'IdTipoCambio'),
    ('ORIGEN', 'CON_Origen', 'IdOrigen'),
    ('CUENTADESTINO', 'CON_CuentaDestinoRegla', 'IdCuentaDestinoRegla'),
    ('ASIENTO', 'CON_Asiento', 'IdAsiento'),
    ('COMPRA', 'COM_Compra', 'IdCompra'),
    ('VENTA', 'VEN_Venta', 'IdVenta'),
    ('CAJABANCO', 'BAN_MovimientoBanco', 'IdMovimientoBanco');

DECLARE @Tabla SYSNAME;
DECLARE @ColumnaId SYSNAME;
DECLARE @Sql NVARCHAR(MAX);

DECLARE cursor_trazabilidad CURSOR LOCAL FAST_FORWARD FOR
    SELECT Tabla, ColumnaId FROM @Tablas;

OPEN cursor_trazabilidad;
FETCH NEXT FROM cursor_trazabilidad INTO @Tabla, @ColumnaId;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF COL_LENGTH(N'dbo.' + @Tabla, 'FechaRegistro') IS NULL
    BEGIN
        SET @Sql = N'ALTER TABLE dbo.' + QUOTENAME(@Tabla) + N' ADD FechaRegistro DATETIME2(0) NULL CONSTRAINT ' + QUOTENAME(N'DF_' + @Tabla + N'_FechaRegistro') + N' DEFAULT (SYSDATETIME());';
        EXEC sys.sp_executesql @Sql;
    END;

    IF COL_LENGTH(N'dbo.' + @Tabla, 'UsuarioRegistro') IS NULL
    BEGIN
        SET @Sql = N'ALTER TABLE dbo.' + QUOTENAME(@Tabla) + N' ADD UsuarioRegistro NVARCHAR(450) NULL;';
        EXEC sys.sp_executesql @Sql;
    END;

    IF COL_LENGTH(N'dbo.' + @Tabla, 'FechaActualizacion') IS NULL
    BEGIN
        SET @Sql = N'ALTER TABLE dbo.' + QUOTENAME(@Tabla) + N' ADD FechaActualizacion DATETIME2(0) NULL;';
        EXEC sys.sp_executesql @Sql;
    END;

    IF COL_LENGTH(N'dbo.' + @Tabla, 'UsuarioActualizacion') IS NULL
    BEGIN
        SET @Sql = N'ALTER TABLE dbo.' + QUOTENAME(@Tabla) + N' ADD UsuarioActualizacion NVARCHAR(450) NULL;';
        EXEC sys.sp_executesql @Sql;
    END;

    FETCH NEXT FROM cursor_trazabilidad INTO @Tabla, @ColumnaId;
END;

CLOSE cursor_trazabilidad;
DEALLOCATE cursor_trazabilidad;
GO
