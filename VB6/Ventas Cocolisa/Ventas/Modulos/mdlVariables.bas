Attribute VB_Name = "mdlVariables"
Option Explicit
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

Public cod_Prov         As String
Public des_prov         As String
Public ruc_prov         As String
Public dir_prov         As String
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

