-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Habilita el origen base de aplicaciones NC para empresas ya existentes.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Inserta o reactiva el origen 47 Aplicaciones N/C en el maestro y en cada empresa para soportar la provision APNC.

MERGE dbo.CON_OrigenMaestro AS destino
USING
(
    VALUES
        ('47', N'APLICACIONES N/C', N'CONTABILIDAD', 1, 180)
) AS fuente (CodigoOrigen, NombreOrigen, ModuloOrigen, PermiteRegistroManual, Orden)
    ON destino.CodigoOrigen = fuente.CodigoOrigen
WHEN MATCHED THEN
    UPDATE
    SET destino.NombreOrigen = fuente.NombreOrigen,
        destino.ModuloOrigen = fuente.ModuloOrigen,
        destino.PermiteRegistroManual = fuente.PermiteRegistroManual,
        destino.Orden = fuente.Orden,
        destino.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoOrigen, NombreOrigen, ModuloOrigen, PermiteRegistroManual, Estado, Orden)
    VALUES (fuente.CodigoOrigen, fuente.NombreOrigen, fuente.ModuloOrigen, fuente.PermiteRegistroManual, 1, fuente.Orden);

MERGE dbo.CON_Origen AS destino
USING
(
    SELECT
        e.IdEmpresa,
        CAST('47' AS VARCHAR(10)) AS CodigoOrigen,
        CAST(N'APLICACIONES N/C' AS NVARCHAR(150)) AS NombreOrigen,
        CAST(N'CONTABILIDAD' AS NVARCHAR(50)) AS ModuloOrigen,
        CAST(1 AS BIT) AS PermiteRegistroManual
    FROM dbo.SEG_Empresa AS e
) AS fuente
    ON destino.IdEmpresa = fuente.IdEmpresa
   AND destino.CodigoOrigen = fuente.CodigoOrigen
WHEN MATCHED THEN
    UPDATE
    SET destino.NombreOrigen = fuente.NombreOrigen,
        destino.ModuloOrigen = fuente.ModuloOrigen,
        destino.PermiteRegistroManual = fuente.PermiteRegistroManual,
        destino.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT
    (
        IdEmpresa,
        CodigoOrigen,
        NombreOrigen,
        ModuloOrigen,
        PermiteRegistroManual,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        fuente.IdEmpresa,
        fuente.CodigoOrigen,
        fuente.NombreOrigen,
        fuente.ModuloOrigen,
        fuente.PermiteRegistroManual,
        1,
        N'codex'
    );
