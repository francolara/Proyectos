-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Inserta monedas base iniciales para Dbsisadm.
-- =============================================

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_Moneda AS m
    WHERE m.CodigoMoneda = 'PEN'
)
BEGIN
    INSERT INTO dbo.ADM_Moneda
    (
        CodigoMoneda,
        NombreMoneda,
        SimboloMoneda,
        EsMonedaBase,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        'PEN',
        N'Sol',
        N'S/',
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_Moneda AS m
    WHERE m.CodigoMoneda = 'USD'
)
BEGIN
    INSERT INTO dbo.ADM_Moneda
    (
        CodigoMoneda,
        NombreMoneda,
        SimboloMoneda,
        EsMonedaBase,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        'USD',
        N'Dolar estadounidense',
        N'US$',
        0,
        1,
        N'codex'
    );
END;
