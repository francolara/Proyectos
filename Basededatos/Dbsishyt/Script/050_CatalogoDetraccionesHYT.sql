-- Firma: FRANCO LARA | Create date: 10/11/2026 | Copia los 37 códigos de FralseCont.dbo.ADM_DetraccionSunat.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
INSERT INTO dbo.DetraccionCatalogoHYT(Codigo,Nombre,Porcentaje,Activo)
SELECT origen.Codigo,origen.Nombre,origen.Porcentaje,origen.Activo
FROM (VALUES
        (N'001', N'Azúcar y melaza de caña', 10.0000, 1),
        (N'002', N'Arroz', 10.0000, 1),
        (N'003', N'Alcohol etílico', 10.0000, 1),
        (N'004', N'Recursos hidrobiológicos', 10.0000, 1),
        (N'005', N'Maíz amarillo duro', 10.0000, 1),
        (N'006', N'Caña de azúcar', 10.0000, 1),
        (N'007', N'Madera', 10.0000, 1),
        (N'008', N'Arena y piedra', 10.0000, 1),
        (N'009', N'Residuos, subproductos, desechos, recortes y desperdicios', 15.0000, 1),
        (N'010', N'Bienes gravados con el IGV o renuncia a la exoneración', 1.5000, 1),
        (N'011', N'Intermediación laboral y tercerización', 12.0000, 1),
        (N'012', N'Animales vivos', 10.0000, 1),
        (N'013', N'Carnes y despojos comestibles', 10.0000, 1),
        (N'014', N'Abonos, cueros y pieles de origen animal', 10.0000, 1),
        (N'015', N'Aceite de pescado', 10.0000, 1),
        (N'016', N'Harina, polvo y pellets de pescado, crustáceos, moluscos y demás invertebrados acuáticos', 10.0000, 1),
        (N'017', N'Arrendamiento de bienes muebles', 10.0000, 1),
        (N'018', N'Mantenimiento y reparación de bienes muebles', 12.0000, 1),
        (N'019', N'Movimiento de carga', 10.0000, 1),
        (N'020', N'Otros servicios empresariales', 12.0000, 1),
        (N'021', N'Leche', 10.0000, 1),
        (N'022', N'Comisión mercantil', 12.0000, 1),
        (N'023', N'Fabricación de bienes por encargo', 12.0000, 1),
        (N'024', N'Servicio de transporte de personas', 10.0000, 1),
        (N'025', N'Servicio de transporte de carga', 4.0000, 1),
        (N'026', N'Transporte de pasajeros', 10.0000, 1),
        (N'027', N'Contratos de construcción', 4.0000, 1),
        (N'028', N'Oro gravado con el IGV', 10.0000, 1),
        (N'029', N'Paprika y otros frutos de los géneros capsicum o pimienta', 10.0000, 1),
        (N'030', N'Minerales metálicos no auríferos', 10.0000, 1),
        (N'031', N'Bienes exonerados del IGV', 1.5000, 1),
        (N'032', N'Oro y demás minerales metálicos exonerados del IGV', 10.0000, 1),
        (N'033', N'Demás servicios gravados con el IGV', 12.0000, 1),
        (N'034', N'Minerales no metálicos', 10.0000, 1),
        (N'035', N'Bien inmueble gravado con IGV', 4.0000, 1),
        (N'036', N'Plomo', 15.0000, 1),
        (N'037', N'Ley 30737', 4.0000, 1)
) AS origen(Codigo,Nombre,Porcentaje,Activo)
WHERE NOT EXISTS(SELECT 1 FROM dbo.DetraccionCatalogoHYT AS actual WHERE actual.Codigo=origen.Codigo);
GO
