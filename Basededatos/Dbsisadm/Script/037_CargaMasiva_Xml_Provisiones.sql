-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Habilita carga masiva XML de compras y ventas, permite cuenta contable pendiente en detalle y ajusta tipos de comprobante para importacion.
-- =============================================

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.COM_CompraDetalle')
      AND name = N'IdPlanCuenta'
      AND is_nullable = 0
)
BEGIN
    ALTER TABLE dbo.COM_CompraDetalle
        ALTER COLUMN IdPlanCuenta INT NULL;
END;

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.VEN_VentaDetalle')
      AND name = N'IdPlanCuenta'
      AND is_nullable = 0
)
BEGIN
    ALTER TABLE dbo.VEN_VentaDetalle
        ALTER COLUMN IdPlanCuenta INT NULL;
END;

IF EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante
    WHERE CodigoTipoComprobante = '03'
)
BEGIN
    UPDATE dbo.ADM_TipoComprobante
    SET UsoCompras = 1,
        UsuarioRegistro = N'FRANCO LARA'
    WHERE CodigoTipoComprobante = '03';
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante
    WHERE CodigoTipoComprobante = '02'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '02',
        N'Recibo por Honorarios',
        1,
        0,
        1,
        N'FRANCO LARA'
    );
END;
