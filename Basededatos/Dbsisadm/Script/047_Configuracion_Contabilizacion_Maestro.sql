-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Registra la configuracion maestra inicial de origenes para los modulos contables bajo el escenario PROVISION.
-- =============================================

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys AS fk
    WHERE fk.name = N'FK_CON_ConfiguracionContabilizacionMaestro_CON_OrigenMaestro'
      AND fk.parent_object_id = OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacionMaestro')
)
BEGIN
    ALTER TABLE dbo.CON_ConfiguracionContabilizacionMaestro
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacionMaestro_CON_OrigenMaestro
            FOREIGN KEY (CodigoOrigen) REFERENCES dbo.CON_OrigenMaestro (CodigoOrigen);
END;

MERGE dbo.CON_ConfiguracionContabilizacionMaestro AS destino
USING
(
    VALUES
        ('COM',  'PROVISION', '44', N'Provision Compras',              1, 1,  10),
        ('VEN',  'PROVISION', '45', N'Provision Ventas',               1, 1,  20),
        ('EGR',  'PROVISION', '02', N'Provision Egresos',              1, 1,  30),
        ('ING',  'PROVISION', '01', N'Provision Ingresos',             1, 1,  40),
        ('APNC', 'PROVISION', '47', N'Provision Aplicaciones',         1, 1,  50),
        ('DET',  'PROVISION', '50', N'Provision Detracciones',         1, 1,  60),
        ('PER',  'PROVISION', '73', N'Provision Percepciones',         1, 1,  70),
        ('DIF',  'PROVISION', '88', N'Provision Diferencia en Cambio', 1, 1,  80),
        ('AJU',  'PROVISION', '67', N'Provision Ajuste de Cuentas',    1, 1,  90),
        ('APR',  'PROVISION', '00', N'Provision Asiento de Apertura',  1, 1, 100),
        ('CIE',  'PROVISION', '77', N'Provision Asiento de Cierre',    1, 1, 110)
) AS fuente
(
    ModuloOperacion,
    EscenarioOperacion,
    CodigoOrigen,
    Descripcion,
    GeneraAsientoAutomatico,
    UsaTipoCambio,
    Orden
)
    ON destino.ModuloOperacion = fuente.ModuloOperacion
   AND destino.EscenarioOperacion = fuente.EscenarioOperacion
WHEN MATCHED THEN
    UPDATE
    SET destino.CodigoOrigen = fuente.CodigoOrigen,
        destino.Descripcion = fuente.Descripcion,
        destino.GeneraAsientoAutomatico = fuente.GeneraAsientoAutomatico,
        destino.UsaTipoCambio = fuente.UsaTipoCambio,
        destino.Activo = 1,
        destino.Orden = fuente.Orden
WHEN NOT MATCHED BY TARGET THEN
    INSERT
    (
        ModuloOperacion,
        EscenarioOperacion,
        CodigoOrigen,
        Descripcion,
        GeneraAsientoAutomatico,
        UsaTipoCambio,
        Activo,
        Orden,
        UsuarioRegistro
    )
    VALUES
    (
        fuente.ModuloOperacion,
        fuente.EscenarioOperacion,
        fuente.CodigoOrigen,
        fuente.Descripcion,
        fuente.GeneraAsientoAutomatico,
        fuente.UsaTipoCambio,
        1,
        fuente.Orden,
        N'sistema'
    );

