-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Agrega maestro general de detracciones, documento pendiente por compra, modulo contable de detracciones y cuenta SPOT para provisiones de compras.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Registra SPOT mediante CodigoCuenta y mantiene el CHECK de modulos alineado con el catalogo contable vigente.

IF COL_LENGTH(N'dbo.COM_Compra', N'TieneDetraccion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD TieneDetraccion BIT NOT NULL CONSTRAINT DF_COM_Compra_TieneDetraccion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'IdDetraccionSunat') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD IdDetraccionSunat INT NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'PorcentajeDetraccion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD PorcentajeDetraccion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajeDetraccion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'ImporteDetraccion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD ImporteDetraccion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImporteDetraccion DEFAULT (0);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_COM_Compra_ADM_DetraccionSunat'
)
AND COL_LENGTH(N'dbo.COM_Compra', N'IdDetraccionSunat') IS NOT NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_DetraccionSunat
            FOREIGN KEY (IdDetraccionSunat) REFERENCES dbo.ADM_DetraccionSunat (IdDetraccionSunat);
END;

IF OBJECT_ID(N'dbo.CK_CON_ConfiguracionContabilizacion_ModuloOperacion', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        DROP CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion;
END;

ALTER TABLE dbo.CON_ConfiguracionContabilizacion
    ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion
        CHECK (ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC', 'DET', 'PER', 'DIF', 'AJU', 'APR', 'CIE'));

MERGE dbo.CON_TipoImpuesto AS target
USING
(
    SELECT CAST('SPOT' AS VARCHAR(10)) AS CodigoSunat, CAST(N'Sistema SPOT (detracción)' AS NVARCHAR(100)) AS NombreImpuesto
) AS source
    ON target.CodigoSunat = source.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        target.NombreImpuesto = source.NombreImpuesto,
        target.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoSunat, NombreImpuesto, CodigoCuenta, Estado)
    VALUES (source.CodigoSunat, source.NombreImpuesto, NULL, 1);

MERGE dbo.ADM_DetraccionSunat AS target
USING
(
    SELECT '001' AS CodigoSunat, N'Azúcar y melaza de caña' AS Descripcion, CAST(10.0000 AS DECIMAL(7,4)) AS Porcentaje UNION ALL
    SELECT '002', N'Arroz', 10.0000 UNION ALL
    SELECT '003', N'Alcohol etílico', 10.0000 UNION ALL
    SELECT '004', N'Recursos hidrobiológicos', 10.0000 UNION ALL
    SELECT '005', N'Maíz amarillo duro', 10.0000 UNION ALL
    SELECT '006', N'Caña de azúcar', 10.0000 UNION ALL
    SELECT '007', N'Madera', 10.0000 UNION ALL
    SELECT '008', N'Arena y piedra', 10.0000 UNION ALL
    SELECT '009', N'Residuos, subproductos, desechos, recortes y desperdicios', 15.0000 UNION ALL
    SELECT '010', N'Bienes gravados con el IGV o renuncia a la exoneración', 1.5000 UNION ALL
    SELECT '011', N'Intermediación laboral y tercerización', 12.0000 UNION ALL
    SELECT '012', N'Animales vivos', 10.0000 UNION ALL
    SELECT '013', N'Carnes y despojos comestibles', 10.0000 UNION ALL
    SELECT '014', N'Abonos, cueros y pieles de origen animal', 10.0000 UNION ALL
    SELECT '015', N'Aceite de pescado', 10.0000 UNION ALL
    SELECT '016', N'Harina, polvo y pellets de pescado, crustáceos, moluscos y demás invertebrados acuáticos', 10.0000 UNION ALL
    SELECT '017', N'Arrendamiento de bienes muebles', 10.0000 UNION ALL
    SELECT '018', N'Mantenimiento y reparación de bienes muebles', 12.0000 UNION ALL
    SELECT '019', N'Movimiento de carga', 10.0000 UNION ALL
    SELECT '020', N'Otros servicios empresariales', 12.0000 UNION ALL
    SELECT '021', N'Leche', 10.0000 UNION ALL
    SELECT '022', N'Comisión mercantil', 12.0000 UNION ALL
    SELECT '023', N'Fabricación de bienes por encargo', 12.0000 UNION ALL
    SELECT '024', N'Servicio de transporte de personas', 10.0000 UNION ALL
    SELECT '025', N'Servicio de transporte de carga', 4.0000 UNION ALL
    SELECT '026', N'Transporte de pasajeros', 10.0000 UNION ALL
    SELECT '027', N'Contratos de construcción', 4.0000 UNION ALL
    SELECT '028', N'Oro gravado con el IGV', 10.0000 UNION ALL
    SELECT '029', N'Paprika y otros frutos de los géneros capsicum o pimienta', 10.0000 UNION ALL
    SELECT '030', N'Minerales metálicos no auríferos', 10.0000 UNION ALL
    SELECT '031', N'Bienes exonerados del IGV', 1.5000 UNION ALL
    SELECT '032', N'Oro y demás minerales metálicos exonerados del IGV', 10.0000 UNION ALL
    SELECT '033', N'Demás servicios gravados con el IGV', 12.0000 UNION ALL
    SELECT '034', N'Minerales no metálicos', 10.0000 UNION ALL
    SELECT '035', N'Bien inmueble gravado con IGV', 4.0000 UNION ALL
    SELECT '036', N'Plomo', 15.0000 UNION ALL
    SELECT '037', N'Ley 30737', 4.0000
) AS source
    ON target.CodigoSunat = source.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        target.Descripcion = source.Descripcion,
        target.Porcentaje = source.Porcentaje,
        target.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoSunat, Descripcion, Porcentaje, Estado)
    VALUES (source.CodigoSunat, source.Descripcion, source.Porcentaje, 1);
