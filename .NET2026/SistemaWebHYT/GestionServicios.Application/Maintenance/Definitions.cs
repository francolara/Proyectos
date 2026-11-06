namespace GestionServicios.Application.Maintenance;

// Firma: FRANCO LARA - 05/11/2026 | Códigos automáticos, abreviaturas, posición en tarifarios y número de documento con tipo oculto, Código SAP, enlaces por cliente y descripción de material de tres líneas.
// Lista cerrada compilada; nunca se aceptan nombres de tabla/SP desde formularios.
public static class MaintenanceDefinitions
{
    public static IReadOnlyList<EntityDefinition> All { get; } =
    [
        new("clientes", "Clientes", "Cliente", "CLIENTES", [
            new("Codigo", "Código interno", "AutoCode", false, 40, null, null, null),
            new("CodigoSAP", "Código SAP", "Text", false, 60, null, null, null),
            new("TipoEntidad", "Tipo de entidad", "Select", true, 20, null, "JURIDICA", "EntityType"),
            new("TipoIdentificacion", "Tipo de documento", "Hidden", false, 20, null, null, null),
            new("NumeroIdentificacion", "Número de documento", "Text", true, 30, null, null, null),
            new("RazonSocial", "Razón social", "Text", false, 200, null, null, null),
            new("Nombres", "Nombres", "Text", false, 150, null, null, null),
            new("Apellidos", "Apellidos", "Text", false, 150, null, null, null),
            new("DireccionFiscal", "Dirección fiscal", "Text", false, 300, null, null, null),
            new("CodigoUbigeo", "Ubigeo", "Ubigeo", false, 6, null, null, null),
            new("UrlCargaDocumentos", "Enlace anterior", "Hidden", false, 2048, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, false),
        new("tiendas", "Tiendas", "Tienda", "TIENDAS", [
            new("ClienteId", "Cliente", "Reference", true, 0, "clientes", null, null),
            new("Codigo", "Código interno", "AutoCode", false, 40, null, null, null),
            new("Nombre", "Nombre de tienda", "Text", true, 200, null, null, null),
            new("Direccion", "Dirección", "Text", false, 300, null, null, null),
            new("CodigoUbigeo", "Ubigeo", "Ubigeo", false, 6, null, null, null),
            new("CentroSAPFICO", "Centro SAP FICO", "Text", false, 50, null, null, null),
            new("ZonaId", "Zona", "Reference", false, 0, "zonas", null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], "ClienteId", false),
        new("contactos", "Contactos", "ContactoCliente", "CLIENTES", [
            new("ClienteId", "Cliente", "Reference", true, 0, "clientes", null, null),
            new("TiendaId", "Tienda (opcional)", "Reference", false, 0, "tiendas", null, null),
            new("Nombre", "Nombre", "Text", true, 150, null, null, null),
            new("Telefono", "Teléfono", "Text", false, 50, null, null, null),
            new("Cargo", "Cargo", "Text", false, 100, null, null, null),
            new("Correo", "Correo electrónico", "Email", false, 256, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], "ClienteId", false),
        new("proveedor-cliente", "Datos de HYT por cliente", "ProveedorCliente", "CLIENTES", [
            new("ClienteId", "Cliente", "Reference", true, 0, "clientes", null, null),
            new("CodigoSAPProveedor", "Código SAP del proveedor", "Text", false, 60, null, null, null),
            new("ContratoMarco", "Contrato marco", "Text", false, 60, null, null, null),
            new("VigenciaDesde", "Vigencia desde", "Date", false, 0, null, null, null),
            new("VigenciaHasta", "Vigencia hasta", "Date", false, 0, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], "ClienteId", false),
        new("unidades", "Unidades de medida", "UnidadMedida", "UNIDADES", [
            new("Codigo", "Abreviatura", "Text", true, 30, null, null, null),
            new("Descripcion", "Descripción", "Text", true, 150, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, false),
        new("zonas", "Zonas", "Zona", "ZONAS", [
            new("Codigo", "Código interno", "AutoCode", false, 30, null, null, null),
            new("Descripcion", "Descripción", "Text", true, 150, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, false),
        new("especialidades", "Especialidades", "Especialidad", "MATERIALES", [
            new("Codigo", "Código interno", "AutoCode", false, 30, null, null, null),
            new("Descripcion", "Descripción", "Text", true, 150, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, false),
        new("tipos-servicio", "Tipos de servicio", "TipoServicio", "MATERIALES", [
            new("Codigo", "Código interno", "AutoCode", false, 30, null, null, null),
            new("Descripcion", "Descripción", "Text", true, 150, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, false),
        new("materiales", "Materiales y servicios", "ServicioMaterial", "MATERIALES", [
            new("CodigoInterno", "Código interno", "AutoCode", false, 60, null, null, null),
            new("CodigoMaterialSAP", "Código de material SAP", "Text", false, 60, null, null, null),
            new("Descripcion", "Descripción", "LongText", true, 1500, null, null, null),
            new("UnidadMedidaId", "Unidad de medida", "Reference", true, 0, "unidades", null, null),
            new("EspecialidadId", "Especialidad", "Reference", false, 0, "especialidades", null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, true),
        new("tarifarios", "Tarifarios", "Tarifario", "TARIFARIOS", [
            new("Codigo", "Código interno", "AutoCode", false, 40, null, null, null),
            new("Nombre", "Nombre", "Text", true, 200, null, null, null),
            new("ClienteId", "Cliente (vacío: genérico)", "Reference", false, 0, "clientes", null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], null, false),
        new("versiones-tarifario", "Versiones de tarifario", "TarifarioVersion", "TARIFARIOS", [
            new("TarifarioId", "Tarifario", "Reference", true, 0, "tarifarios", null, null),
            new("NumeroVersion", "Número de versión", "ReadOnly", false, 0, null, null, null),
            new("VigenciaDesde", "Vigencia desde", "Date", false, 0, null, null, null),
            new("VigenciaHasta", "Vigencia hasta", "Date", false, 0, null, null, null),
            new("CodigoMonedaPredeterminada", "Moneda predeterminada", "Select", true, 3, null, "PEN", "Currency"),
            new("CodigoTipoIGVPredeterminado", "Tratamiento IGV predeterminado", "Select", true, 10, null, null, "TaxTreatment"),
            new("Estado", "Estado", "ReadOnly", false, 20, null, "BORRADOR", null)
        ], "TarifarioId", true),
        new("partidas-tarifario", "Partidas del tarifario", "TarifarioPartida", "TARIFARIOS", [
            new("VersionId", "Versión de tarifario", "Reference", true, 0, "versiones-tarifario", null, null),
            new("MaterialId", "Material / servicio", "Reference", true, 0, "materiales", null, null),
            new("Posicion", "Posición", "Text", false, 60, null, null, null),
            new("CodigoMaterialSAP", "Código de material SAP", "Text", false, 60, null, null, null),
            new("TipoServicioId", "Tipo de servicio", "Reference", false, 0, "tipos-servicio", null, null),
            new("NumeroNecesidad", "Número de necesidad", "Text", false, 60, null, null, null),
            new("CodigoMoneda", "Moneda de referencia", "Select", true, 3, null, "PEN", "Currency"),
            new("CodigoTipoIGV", "Tratamiento IGV de los precios", "Select", true, 10, null, null, "TaxTreatment"),
            new("PrecioPEN", "Precio en soles (PEN)", "Decimal", false, 0, null, null, null),
            new("PrecioUSD", "Precio en dólares (USD), opcional", "Decimal", false, 0, null, null, null),
            new("Activo", "Activo", "Boolean", true, 0, null, "true", null)
        ], "VersionId", true)
    ];
    public static EntityDefinition? Find(string key) => All.FirstOrDefault(item => item.Key == key);
}
