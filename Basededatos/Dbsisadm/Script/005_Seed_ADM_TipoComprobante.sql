-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Inserta tipos de comprobante SUNAT base para compras y ventas.
-- =============================================

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '01'
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
        '01',
        N'Factura',
        1,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '03'
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
        '03',
        N'Boleta de venta',
        0,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '07'
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
        '07',
        N'Nota de credito',
        1,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '08'
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
        '08',
        N'Nota de debito',
        1,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '14'
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
        '14',
        N'Recibo por servicios publicos',
        1,
        0,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '50'
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
        '50',
        N'Declaracion unica de aduanas',
        1,
        0,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '91'
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
        '91',
        N'Comprobante de no domiciliado',
        1,
        0,
        1,
        N'codex'
    );
END;
