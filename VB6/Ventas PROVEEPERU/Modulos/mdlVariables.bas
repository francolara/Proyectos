Attribute VB_Name = "mdlVariables"
Option Explicit
' Firma Codex 15/10/2026: cache de configuracion de FrmBusqueda cargada durante el login.
' Firma Codex 17/10/2026: cache de parametros usados por ordenes de compra, su ayuda de productos y mantenimientos.
Public Cn  As New ADODB.Connection
Public CnConta  As New ADODB.Connection
Public Cnf As New ADODB.Connection
Public strcn    As String
Public gbservidor   As String
Public gbDatabase   As String
Public gbusuario    As String
Public gbPassword   As String
Public gbRutaProductos   As String

Public wingreso     As Integer
Public strRepetirProductosGrid As String
Public gStrServidor             As String
Public gStrMotorBD              As Integer
Public gStrBD                   As String
Public gStrTC                   As Integer
Public gStrUsuario              As String
Public gStrClave                As String
Public gStrRutaRpts             As String
Public intento                  As Byte
Public strReporte               As String
Public strCodAyuda              As String
Public indRespuesta             As Integer
Public valdscto                 As Double
Public valcant                  As Integer

Public Const CONF_CONEXION = "Conexion"
Public Const CONF_OTROS = "Otros"
Public Const CONF_SERVIDOR = "Servidor"
Public Const CONF_MOTORBD = "MotorBD"
Public Const CONF_BD = "BaseDatos"
Public Const CONF_TC = "TipoConexion"
Public Const CONF_USUARIO = "Usuario"
Public Const CONF_CLAVE = "Clave"
Public Const CONF_RUTA_RPTS = "RutaReportes"
Public Const CONF_INDEX_EMPRESA = "Empresa"

Public csql As String
Public glsUser As String
Public glsTC As Double
Public glsEmpresa As String
Public glsPersonaEmpresa As String
Public glsSucursal As String
Public glsNumNiveles As Integer
Public glsSistemaAccess As String
Public glsCodPeriodoINV As String

Public indAdmin As Boolean
Public indCompraUsuario As Boolean
Public strArregloMes(1 To 12) As String

Public glsIGV As Double
Public glsAlmVentas As String
Public glsMonVentas As String
Public glsListaVentas As String
Public glsDecimalesCaja As Integer
Public glsDecimalesTC As Integer
Public glsDecimalesPrecios As Integer
Public glsValidaStock As Boolean
Public glsModVendCampo As Boolean
Public glsSystem As String
Public glsClienteVentas As String
Public glsFormaPagoVentas As String
Public glsRutaImagenProd As String


Public STR_VISUALIZA_OC_FACTURA_PEDIDO As String
Public STR_IMPORTAR_DOCUMENTOS_ENTRE_EMPRESAS As String
Public STR_EMPRESA_SUCURSAL_DOCUMENTOS_ENTRE_EMPRESAS As String
Public STR_BLOQUEA_CHK_AFECTO As String
Public STR_VISUALIZAR_AYUDA_REQUERIMIENTO_COMPRA As String
Public STR_LONGITUD_IMPRESION_DETALLE_PRODUCTO As String
Public STR_DESCRIPCION_AREA_O_UPP As String

Public STR_IMPORTA_ATENCIONES As String
Public STR_APRUEBA_PEDIDO_AUTOMATICO As String
Public STR_VENTA_ELECTRONICA As String

Public STR_STOCK_POR_LOTE As String
Public STR_LIQUIDACIONES As String
Public STR_VALIDA_SEPARACION As String
Public STR_IGV As String
Public STR_IGV_ANT As String
Public STR_PERIODO_CAMBIO_IGV As String
Public STR_CLIENTE_ANULA As String
Public STR_SUCURSAL_PARTIDA_UBIGEO As String
Public STR_SUCURSAL_PARTIDA_DIRECCION As String

'Parametros de frmDocVentas cargados al iniciar sesion.
Public STR_CONCEPTO_VENTA_GRATUITA As String
Public STR_FORMA_PAGOS_X_CLIENTE As String
Public STR_VISUALIZA_CLIENTESXVENDEDOR As String
Public STR_VISUALIZA_NOMBRE_COMERCIAL As String
Public STR_IND_VISUALIZACLIENTES_VENDEDOR As String
Public STR_VIZUALIZA_SON As String
Public STR_VIZUALIZA_CODIGO_RAPIDO As String
Public STR_CONTABILIDAD_ONLINE As String
Public STR_CONTABILIDAD_ONLINE_VB As String
Public STR_MODIFICA_REGDOC_DESPROD As String
Public STR_SOLOVALIDADCTOMAX As String
Public STR_BLOQUEA_DETALLE As String
Public STR_VALIDA_REFERENCIA As String
Public STR_PRODUCTO_ANULA As String
Public STR_EVALUA_ITEM_DETALLE_PRODUCTO As String
Public STR_AGRUPAPRODUCTOS As String
Public STR_IMP_FORMATO_PEDIDO As String
Public STR_CONCEPTO_AUTOMATICO_SALIDA As String
Public STR_ALMACEN_AUTOMATICO_DESTINO As String
Public STR_CONCEPTO_AUTOMATICO_INGRESO As String
Public STR_CODIGO_PRODUCTO_APIMAS As String
Public STR_PORCENTAJE_FACTURACION As String
Public STR_MONEDA_LINEA_CREDITO As String
Public STR_EXTRAE_PRECIO_COMPRAS As String
Public STR_EXTRAE_PRECIO_ULTIMA_VENTA As String
Public STR_FACTURACION_ENTRE_EMPRESAS As String
Public STR_DNSCONTABILIDAD As String
Public STR_TIPO_IGV As String
Public STR_AYUDA_CENTRO_COSTO_DETALLE As String
Public STR_VALIDA_FACTURADO_HCOSTO As String
Public STR_CARPETA_XML_VE As String
Public STR_VALIDA_PERMISOS As String
Public STR_PORCENTAJE_IVAP As String
Public STR_FORMAPAGO_PEDIDO As String
Public STR_CODIGO_SERVICIO_ATENCIONES As String
Public STR_FORMA_PAGOS_ACTIVO_FACTURA As String
Public STR_IMP_COT_ALSISAC As String
Public STR_TRANS_CONTA_CTA_40 As String
Public STR_TRANS_CONTA_CTA_70 As String
Public STR_IVAP_CUENTA As String

'Parametros de frmDocVentas_OC y su ayuda cargados al iniciar sesion.
Public STR_NO_VALIDAR_CONTACTO_OC As String
Public STR_VISUALIZA_CENTROCOSTO_OC As String
Public STR_VALIDAPRESUPUESTO As String
Public STR_VALIDA_HC_OS As String
Public STR_FORMATO_OC As String
Public STR_PROVEEDOR_POR_DEFECTO As String
Public STR_ALMACEN_COMPRAS As String
Public STR_VALIDA_APROBACION_BANDEJA As String
Public STR_VALIDA_PRODUCTOPROVEEDOR_OC As String
Public STR_OPCION_AYUDA_PRODUCTOS As String

'Parametros de mantenimientos cargados al iniciar sesion.
Public STR_CLIENTE_ESPECIAL As String
Public STR_VISUALIZA_LISTA_PRECIO_ALMACEN As String
Public STR_ACTUALIZA_DESCUENTO_NIVEL As String
Public STR_NIVEL_CUENTA_CONTABLE As String
Public STR_VISUALIZA_STOCK_MINIMO As String
Public STR_VISUALIZA_REPORTE_GL As String
Public STR_VALIDA_CODIGO_FABRICANTE As String
Public STR_MUESTRA_ALM_DEFECTO As String
Public STR_TIPO_CAMBIO_BANCOS As String
Public STR_MONTO_MAXIMO_TIPO_DE_CAMBIO As String
Public STR_VISUALIZA_TIPOVENTA As String

Public cod_Prov         As String
Public des_prov         As String
Public ruc_prov         As String
Public dir_prov         As String
Public Ubigeo_prov         As String
Public wcodruc_Nuevo    As String
Public wcod_obliga      As String
Public wdes_obliga      As String

Public glsSoloGuiaMueveStock As Boolean
Public glsRecepcionAuto As Boolean
Public glsDctoMinValidacion As Double
Public glsRuta_Access       As String
Public glsOrigen_Contable   As String
Public glsCuenta_Igv_Ventas As String
Public glsRuta_Access_Conta As String
Public sw_proveedor     As Boolean
Public glsGraba_Contado     As String
Public glsTipoCambio        As String
Public glsGrabaTodo         As String
Public glsLeeCodigoBarras   As String
Public glsDctoMinMonto      As Double
Public glsVisualizaCodFab   As String

Public glsEnterAyudaClientes    As Boolean
Public glsEnterAyudaProductos   As Boolean
Public sw_graba_factura  As Boolean 'activar en caso se quiera grabar guia y factura al mismo tiempo
Public strNumDocReferencia As String

Public StrcodSistema            As String
Public glsDsctoConClave As String
Public glsModificarPrecio As String
Public glsPorcentajeRetencion As Double
Public glsFormatoImpLetra As String
Public glsMotivoSalida As String
Public glsGrabaGuiaFactura  As String
Public GlsVisualiza_Filtro_Documento  As String
Public glsobservacionventas As String
Public glsobservacionventas2 As String

Public rsAsientosContables      As New ADODB.Recordset
Public strcnConta    As String

Public stridseriedocREF As String
Public strGlsDocREF     As String
Public stridNumDocREF   As String
Public indEvaluaVacio   As Boolean

Public glsobservacioncliente As String
Public sw_estadistico   As Boolean
Public wusuario         As String

Public strAno          As String
Public strMes          As String

