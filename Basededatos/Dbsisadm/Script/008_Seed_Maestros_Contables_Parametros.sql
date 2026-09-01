-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Precarga inicial de maestros internos para parametros, origenes, plan de cuentas y cuentas destino con ColBalance, moneda y tipo de cambio.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Agrega los parametros maestros CTADETRACCION y CTADEPERCEPCION para configurar cuentas contables de detracciones y percepciones desde ADM_ParametroEmpresa.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Agrega el origen maestro 73 para percepciones de compras y el parametro contable CTADEPERCEPCION.
-- =============================================
-- Firma: FRANCO LARA - 01/07/2026 | Agrega el parametro maestro ORIGEN_DIFERENCIA_CAMBIO como referencia legacy y mantiene disponible el origen 88 para configurar diferencia en cambio desde el modulo web.
-- Firma: FRANCO LARA - 02/07/2026 | Agrega GeneraDiferenciaPorAnalisis al plan de cuentas maestro para heredar la configuracion base o de empresa origen y suma los parametros/origen base de los nuevos procesos AJU, APR y CIE.
-- Firma: FRANCO LARA - 25/08/2026 | Elimina el ejercicio de las cuentas destino maestras; el ejercicio se asigna al materializar la regla por empresa.
-- Firma: FRANCO LARA - 27/08/2026 | Corrige la carga del plan maestro para no actualizar ni insertar la columna eliminada GeneraDiferenciaPorAnalisis.

MERGE dbo.ADM_ParametroMaestro AS destino
USING
(
    VALUES
        ('CONTABLE', 'BALANCE_COMPROBACION_FORMATO2', N'N', N'S: Para Ideas', NULL, NULL, 10),
        ('CONTABLE', 'CUENTAGANANCIA', N'77611009', N'Cuenta Ganancia para Ajustes', NULL, NULL, 20),
        ('CONTABLE', 'CUENTAGANANCIA_DC', N'77611009', N'Cuenta Ganancia para Diferencia en Cambio', NULL, NULL, 30),
        ('CONTABLE', 'CUENTAGANANCIA_AJ', N'77611009', N'Cuenta Ganancia para Ajuste de Cuentas', NULL, NULL, 35),
        ('CONTABLE', 'CUENTAPERDIDA', N'97611009', N'Cuenta Perdida para Ajustes', NULL, NULL, 40),
        ('CONTABLE', 'CUENTAPERDIDA_DC', N'97611009', N'Cuenta Perdida para Diferencia en Cambio', NULL, NULL, 50),
        ('CONTABLE', 'CUENTAPERDIDA_AJ', N'97611009', N'Cuenta Perdida para Ajuste de Cuentas', NULL, NULL, 45),
        ('CONTABLE', 'ORIGEN_DIFERENCIA_CAMBIO', N'88', N'Origen sugerido para el proceso de diferencia en cambio', NULL, NULL, 55),
        ('CONTABLE', 'ORIGEN_ASIENTO_APERTURA', N'00', N'Origen sugerido para el proceso de asiento de apertura', NULL, NULL, 56),
        ('CONTABLE', 'ORIGEN_AJUSTE_CUENTAS', N'67', N'Origen sugerido para el proceso de ajuste de cuentas', NULL, NULL, 57),
        ('CONTABLE', 'ORIGEN_ASIENTO_CIERRE', N'77', N'Origen sugerido para el proceso de asiento de cierre', NULL, NULL, 58),
        ('CONTABLE', 'FORMATO_VOUCHER', N'', N'Vacio por defecto, 1: Ideas', NULL, NULL, 60),
        ('CONTABLE', 'GRADO1_LONG', N'2', N'Indicar la longitud del grado 1', NULL, NULL, 70),
        ('CONTABLE', 'GRADO2_LONG', N'1', N'Indicar la longitud del grado 2', NULL, NULL, 80),
        ('CONTABLE', 'GRADO3_LONG', N'1', N'Indicar la longitud del grado 3', NULL, NULL, 90),
        ('CONTABLE', 'GRADO4_LONG', N'2', N'Indicar la longitud del grado 4', NULL, NULL, 100),
        ('CONTABLE', 'GRADO5_LONG', N'2', N'Indicar la longitud del grado 5', NULL, NULL, 110),
        ('CONTABLE', 'GRADO_MAXIMO', N'5', N'Grado maximo del plan contable', NULL, NULL, 120),
        ('CONTABLE', 'IMPRESION_A4', N'1', N'Usar formato A4 en reportes contables', NULL, NULL, 130),
        ('CONTABLE', 'TIPO_CAMBIO_SBS_CIERRE', N'', N'S: trabajar con tipo de cambio SBS en cierres y diferencia de cambio', NULL, NULL, 140),
        ('ADMINISTRATIVO', 'CONTABILIDAD_ONLINE', N'0', N'Indica si la empresa trabaja con contabilidad online', NULL, NULL, 210),
        ('ADMINISTRATIVO', 'CONTABILIDAD_ONLINE_VB', N'0', N'1: requiere aprobacion del contador para modificar provision', NULL, NULL, 220),
        ('ADMINISTRATIVO', 'CLIENTEVENTAS', N'', N'Cliente por defecto para boletas y tickets', NULL, NULL, 230),
        ('ADMINISTRATIVO', 'CALCULA_VCTO_COMPRAS', N'RECEPCION', N'Fecha base para calcular vencimiento en compras', NULL, NULL, 240),
        ('ADMINISTRATIVO', 'CTARETENCION', N'40172000', N'Cuenta de retencion', NULL, NULL, 250),
        ('ADMINISTRATIVO', 'CTA_DEBE_CONSUMO', N'92111001', N'Cuenta debe para consumo', NULL, NULL, 260),
        ('ADMINISTRATIVO', 'CTA_HABER_CONSUMO', N'79111001', N'Cuenta haber para consumo', NULL, NULL, 270),
        ('ADMINISTRATIVO', 'CTADETRACCION', N'', N'Cuenta contable para el Sistema SPOT (detraccion)', NULL, NULL, 280),
        ('ADMINISTRATIVO', 'CTADEPERCEPCION', N'', N'Cuenta contable para percepciones de compras', NULL, NULL, 285),
        ('COMPRAS', 'ORIGEN_REGISTRO_COMPRAS', N'44', N'Origen contable para registro de compras', NULL, NULL, 310),
        ('VENTAS', 'ORIGEN_REGISTRO_VENTAS', N'45', N'Origen contable para registro de ventas', NULL, NULL, 320)
) AS fuente (TipoParametro, CodigoParametro, ValorParametro, DescripcionParametro, FecIni, FecFin, Orden)
    ON destino.TipoParametro = fuente.TipoParametro
   AND destino.CodigoParametro = fuente.CodigoParametro
WHEN MATCHED THEN
    UPDATE
    SET destino.ValorParametro = fuente.ValorParametro,
        destino.DescripcionParametro = fuente.DescripcionParametro,
        destino.FecIni = fuente.FecIni,
        destino.FecFin = fuente.FecFin,
        destino.Orden = fuente.Orden,
        destino.Activo = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (TipoParametro, CodigoParametro, ValorParametro, DescripcionParametro, FecIni, FecFin, Orden, Activo)
    VALUES (fuente.TipoParametro, fuente.CodigoParametro, fuente.ValorParametro, fuente.DescripcionParametro, fuente.FecIni, fuente.FecFin, fuente.Orden, 1);

MERGE dbo.CON_OrigenMaestro AS destino
USING
(
    VALUES
        ('00', N'ASIENTO DE APERTURA', N'CONTABILIDAD', 1, 10),
        ('01', N'INGRESOS', N'TESORERIA', 1, 20),
        ('02', N'EGRESOS', N'TESORERIA', 1, 30),
        ('03', N'CUENTAS DE ORDEN', N'CONTABILIDAD', 1, 40),
        ('04', N'PROVEEDORES', N'COMPRAS', 1, 50),
        ('06', N'CLIENTES', N'VENTAS', 1, 60),
        ('07', N'TRIBUTOS POR PAGAR', N'CONTABILIDAD', 1, 70),
        ('08', N'PROVISIONES', N'CONTABILIDAD', 1, 80),
        ('09', N'REGULARIZACIONES', N'CONTABILIDAD', 1, 90),
        ('10', N'PLANILLA', N'RRHH', 1, 100),
        ('11', N'PLANILLA DE CONSTRUCCION CIVIL', N'RRHH', 1, 110),
        ('15', N'INGRESOS DIFERIDOS', N'VENTAS', 1, 120),
        ('16', N'EGRESOS DIFERIDOS', N'COMPRAS', 1, 130),
        ('20', N'INVENTARIOS', N'INVENTARIO', 1, 140),
        ('44', N'REGISTRO DE COMPRAS', N'COMPRAS', 0, 150),
        ('45', N'REGISTRO DE VENTAS', N'VENTAS', 0, 160),
        ('46', N'MATERIA PRIMA POR RECIBIR', N'COMPRAS', 1, 170),
        ('47', N'APLICACIONES N/C', N'CONTABILIDAD', 1, 180),
        ('48', N'APLICACIONES CLIENTES', N'VENTAS', 1, 190),
        ('49', N'VARIOS', N'CONTABILIDAD', 1, 200),
        ('50', N'DETRACCIONES', N'TESORERIA', 1, 210),
        ('66', N'AJUSTE CORRECCION MONETARIA', N'CONTABILIDAD', 1, 220),
        ('67', N'AJUSTE DE CUENTAS', N'CONTABILIDAD', 1, 225),
        ('70', N'DEPRECIACION', N'ACTIVOS', 1, 230),
        ('71', N'FACTURAS POR COBRAR INGRESOS', N'VENTAS', 1, 240),
        ('72', N'PERCEPCIONES', N'VENTAS', 1, 250),
        ('73', N'PERCEPCIONES COMPRAS', N'COMPRAS', 1, 255),
        ('74', N'REGISTRO DE COMPRAS EXTORNO DUAS', N'COMPRAS', 0, 260),
        ('77', N'ASIENTO DE CIERRE', N'CONTABILIDAD', 1, 270),
        ('78', N'COSTO PRODUCTIVO', N'COSTOS', 1, 280),
        ('79', N'REGULARIZACIONES', N'CONTABILIDAD', 1, 290),
        ('80', N'CANJES CLIENTES-PROVEEDORES', N'CONTABILIDAD', 1, 300),
        ('81', N'COSTO PRODUCTIVO', N'COSTOS', 1, 310),
        ('83', N'LETRAS POR PAGAR', N'TESORERIA', 1, 320),
        ('84', N'REGISTRO COMPRAS', N'COMPRAS', 0, 330),
        ('85', N'AJUSTES CENTIMOS', N'CONTABILIDAD', 1, 340),
        ('86', N'COSTO PRODUCTIVO', N'COSTOS', 1, 350),
        ('87', N'REGISTRO DE VENTAS - LIMA', N'VENTAS', 0, 360),
        ('88', N'AJUSTE POR D/C', N'CONTABILIDAD', 1, 370),
        ('89', N'AJUSTE POR D/C - SOLES', N'CONTABILIDAD', 1, 380),
        ('90', N'PLANILLA - OBREROS', N'RRHH', 1, 390),
        ('91', N'PROVISION RRHH', N'RRHH', 1, 400),
        ('92', N'PROVISIONES PRESTAMOS INTERCOMPANIA', N'CONTABILIDAD', 1, 410),
        ('93', N'LIQUIDACION BENEF. SOC. - EMPLEADOS', N'RRHH', 1, 420),
        ('94', N'LIQUIDACION BENEF. SOC. - OBREROS', N'RRHH', 1, 430),
        ('95', N'LETRAS DESCTO - LIMA', N'TESORERIA', 1, 440),
        ('96', N'PLLA SUELDOS', N'RRHH', 1, 450),
        ('97', N'PLLA SALARIOS', N'RRHH', 1, 460),
        ('98', N'ASIENTOS DE DIARIO', N'CONTABILIDAD', 1, 470),
        ('99', N'COSTO PRODUCTIVO', N'COSTOS', 1, 480)
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

MERGE dbo.CON_PlanCuentaMaestro AS destino
USING
(
    VALUES
        ('10', NULL, N'EFECTIVO Y EQUIVALENTES DE EFECTIVO', 1, 'I', '', '', 0, 0, 10),
        ('101', '10', N'CAJA', 2, 'I', '', '', 0, 0, 20),
        ('1011', '101', N'EFECTIVO', 3, 'I', '', '', 0, 0, 30),
        ('10111001', '1011', N'EFECTIVO MN', 4, 'I', 'PEN', '', 1, 0, 40),
        ('10111002', '1011', N'EFECTIVO ME', 4, 'I', 'USD', 'V', 1, 0, 50),
        ('12', NULL, N'CUENTAS POR COBRAR COMERCIALES - TERCEROS', 1, 'I', '', '', 0, 0, 60),
        ('121', '12', N'FACTURAS, BOLETAS Y OTROS COMPROBANTES POR COBRAR', 2, 'I', '', '', 0, 0, 70),
        ('1212', '121', N'EMITIDAS EN CARTERA', 3, 'I', '', '', 0, 0, 80),
        ('12121001', '1212', N'CLIENTES MN', 4, 'I', 'PEN', '', 1, 0, 90),
        ('40', NULL, N'TRIBUTOS, CONTRAPRESTACIONES Y APORTES AL SISTEMA PUBLICO', 1, 'I', '', '', 0, 0, 100),
        ('401', '40', N'GOBIERNO NACIONAL', 2, 'I', '', '', 0, 0, 110),
        ('4011', '401', N'IMPUESTO GENERAL A LAS VENTAS', 3, 'I', '', '', 0, 0, 120),
        ('40111001', '4011', N'IGV CUENTA PROPIA', 4, 'I', 'PEN', '', 1, 0, 130),
        ('42', NULL, N'CUENTAS POR PAGAR COMERCIALES - TERCEROS', 1, 'I', '', '', 0, 0, 140),
        ('421', '42', N'FACTURAS, BOLETAS Y OTROS COMPROBANTES POR PAGAR', 2, 'I', '', '', 0, 0, 150),
        ('4212', '421', N'EMITIDAS', 3, 'I', '', '', 0, 0, 160),
        ('42121001', '4212', N'PROVEEDORES MN', 4, 'I', 'PEN', '', 1, 0, 170),
        ('60', NULL, N'COMPRAS', 1, 'N', '', '', 0, 0, 180),
        ('601', '60', N'MERCADERIAS', 2, 'N', '', '', 0, 0, 190),
        ('6011', '601', N'MERCADERIAS MANUFACTURADAS', 3, 'N', '', '', 0, 0, 200),
        ('60111001', '6011', N'MERCADERIAS MANUFACTURADAS MN', 4, 'N', 'PEN', '', 1, 0, 210),
        ('63', NULL, N'GASTOS DE SERVICIOS PRESTADOS POR TERCEROS', 1, 'N', '', '', 0, 0, 220),
        ('631', '63', N'TRANSPORTE, CORREOS Y GASTOS DE VIAJE', 2, 'N', '', '', 0, 0, 230),
        ('63111001', '631', N'TRANSPORTE MN', 3, 'N', 'PEN', '', 1, 1, 240),
        ('70', NULL, N'VENTAS', 1, 'R', '', '', 0, 0, 250),
        ('701', '70', N'MERCADERIAS', 2, 'R', '', '', 0, 0, 260),
        ('7011', '701', N'MERCADERIAS MANUFACTURADAS', 3, 'R', '', '', 0, 0, 270),
        ('70111001', '7011', N'VENTA DE MERCADERIAS MN', 4, 'R', 'PEN', '', 1, 0, 280),
        ('79', NULL, N'CARGAS IMPUTABLES A CUENTAS DE COSTOS Y GASTOS', 1, 'S', '', '', 0, 0, 290),
        ('791', '79', N'CARGAS IMPUTABLES A CUENTAS DE COSTOS Y GASTOS', 2, 'S', '', '', 0, 0, 300),
        ('7911', '791', N'CARGAS IMPUTABLES', 3, 'S', '', '', 0, 0, 310),
        ('79111001', '7911', N'CARGAS IMPUTABLES A COSTOS', 4, 'S', 'PEN', '', 1, 0, 320),
        ('79111002', '7911', N'CARGAS IMPUTABLES A GASTOS', 4, 'S', 'PEN', '', 1, 0, 330),
        ('90', NULL, N'COSTOS DE PRODUCCION', 1, 'F', '', '', 0, 0, 340),
        ('901', '90', N'COSTOS POR DISTRIBUIR', 2, 'F', '', '', 0, 0, 350),
        ('9011', '901', N'COSTOS POR DISTRIBUIR', 3, 'F', '', '', 0, 0, 360),
        ('90111001', '9011', N'COSTOS DE MERCADERIA', 4, 'F', 'PEN', '', 1, 1, 370)
) AS fuente (CodigoCuenta, CodigoCuentaPadre, NombreCuenta, NivelCuenta, ColBalance, IdMoneda, TipoCambio, AceptaMovimiento, RequiereCentroCosto, Orden)
    ON destino.CodigoCuenta = fuente.CodigoCuenta
WHEN MATCHED THEN
    UPDATE
    SET destino.CodigoCuentaPadre = fuente.CodigoCuentaPadre,
        destino.NombreCuenta = fuente.NombreCuenta,
        destino.NivelCuenta = fuente.NivelCuenta,
        destino.ColBalance = fuente.ColBalance,
        destino.IdMoneda = fuente.IdMoneda,
        destino.TipoCambio = fuente.TipoCambio,
        destino.AceptaMovimiento = fuente.AceptaMovimiento,
        destino.RequiereCentroCosto = fuente.RequiereCentroCosto,
        destino.Orden = fuente.Orden,
        destino.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoCuenta, CodigoCuentaPadre, NombreCuenta, NivelCuenta, ColBalance, IdMoneda, TipoCambio, AceptaMovimiento, RequiereCentroCosto, Estado, Orden)
    VALUES (fuente.CodigoCuenta, fuente.CodigoCuentaPadre, fuente.NombreCuenta, fuente.NivelCuenta, fuente.ColBalance, fuente.IdMoneda, fuente.TipoCambio, fuente.AceptaMovimiento, fuente.RequiereCentroCosto, 1, fuente.Orden);

MERGE dbo.CON_CuentaDestinoReglaMaestro AS destino
USING
(
    VALUES
        ('90111001', N'Regla base de destino para costo de mercaderia')
) AS fuente (CodigoCuentaOrigen, Observacion)
    ON destino.CodigoCuentaOrigen = fuente.CodigoCuentaOrigen
WHEN MATCHED THEN
    UPDATE
    SET destino.Observacion = fuente.Observacion,
        destino.Activo = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoCuentaOrigen, Activo, Observacion)
    VALUES (fuente.CodigoCuentaOrigen, 1, fuente.Observacion);

MERGE dbo.CON_CuentaDestinoReglaDetalleMaestro AS destino
USING
(
    SELECT
        rm.IdCuentaDestinoReglaMaestro,
        CAST(1 AS SMALLINT) AS Orden,
        CAST('63111001' AS VARCHAR(20)) AS CodigoCuentaDestinoCargo,
        CAST('79111002' AS VARCHAR(20)) AS CodigoCuentaDestinoAbono,
        CAST(100 AS DECIMAL(7,4)) AS Porcentaje
    FROM dbo.CON_CuentaDestinoReglaMaestro AS rm
    WHERE rm.CodigoCuentaOrigen = '90111001'
) AS fuente
    ON destino.IdCuentaDestinoReglaMaestro = fuente.IdCuentaDestinoReglaMaestro
   AND destino.Orden = fuente.Orden
WHEN MATCHED THEN
    UPDATE
    SET destino.CodigoCuentaDestinoCargo = fuente.CodigoCuentaDestinoCargo,
        destino.CodigoCuentaDestinoAbono = fuente.CodigoCuentaDestinoAbono,
        destino.Porcentaje = fuente.Porcentaje,
        destino.Activo = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (IdCuentaDestinoReglaMaestro, Orden, CodigoCuentaDestinoCargo, CodigoCuentaDestinoAbono, Porcentaje, Activo)
    VALUES (fuente.IdCuentaDestinoReglaMaestro, fuente.Orden, fuente.CodigoCuentaDestinoCargo, fuente.CodigoCuentaDestinoAbono, fuente.Porcentaje, 1);
