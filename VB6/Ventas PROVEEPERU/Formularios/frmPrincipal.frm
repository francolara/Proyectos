VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Modulo Ventas - Menu Principal"
   ClientHeight    =   7365
   ClientLeft      =   4650
   ClientTop       =   3585
   ClientWidth     =   13740
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
      Align           =   3  'Align Left
      Height          =   6990
      Left            =   0
      OleObjectBlob   =   "frmPrincipal.frx":27A2
      TabIndex        =   1
      Top             =   0
      Width           =   1485
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6990
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7011
            Picture         =   "frmPrincipal.frx":18274
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7011
            Picture         =   "frmPrincipal.frx":18C6E
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7011
            Picture         =   "frmPrincipal.frx":18F04
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1929E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":19638
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":19A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":19E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1A1BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1A558
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1A8F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1AC8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1B026
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1B3C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1B75A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":1C41C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu07 
      Caption         =   "Ventas"
      Begin VB.Menu mnu0701 
         Caption         =   "Factura"
      End
      Begin VB.Menu mnu0702 
         Caption         =   "Boleta de Venta"
      End
      Begin VB.Menu mnu0703 
         Caption         =   "Pedido"
      End
      Begin VB.Menu mnu0705 
         Caption         =   "Guia de Remision"
      End
      Begin VB.Menu mnu0706 
         Caption         =   "Nota de Credito"
      End
      Begin VB.Menu mnu0707 
         Caption         =   "Nota de Debito"
      End
      Begin VB.Menu mnu0708 
         Caption         =   "Ticket"
      End
      Begin VB.Menu mnu0710 
         Caption         =   "Nota de Venta"
      End
      Begin VB.Menu mnu0709 
         Caption         =   "Bandeja de Pedidos"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu0711 
         Caption         =   "Separacion de Mercaderia"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu0712 
         Caption         =   "Cotizacion"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu0713 
         Caption         =   "Nota de Descuento"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu0715 
         Caption         =   "Guías Madres"
      End
      Begin VB.Menu mnu0716 
         Caption         =   "Liquidaciones de Ventas"
      End
      Begin VB.Menu mnu0717 
         Caption         =   "Atribuciones"
      End
      Begin VB.Menu mnu0718 
         Caption         =   "Atribuciones - Notas de Crédito"
      End
      Begin VB.Menu mnu0714 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu10 
      Caption         =   "Inventario"
      Begin VB.Menu mnu1001 
         Caption         =   "Vale de Ingreso"
      End
      Begin VB.Menu mnu1002 
         Caption         =   "Vale de Salida"
      End
      Begin VB.Menu mnu1003 
         Caption         =   "Transferencia entre almacenes"
      End
      Begin VB.Menu mnu1004 
         Caption         =   "Periodo de Inventario"
      End
      Begin VB.Menu mnu1005 
         Caption         =   "Procesar costos"
      End
      Begin VB.Menu mnu1006 
         Caption         =   "Procesar Saldos"
      End
      Begin VB.Menu mnu1007 
         Caption         =   "Orden de Compra"
      End
      Begin VB.Menu mnu1008 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu02 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu mnu0201 
         Caption         =   "Entidades"
      End
      Begin VB.Menu mnu0209 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnu0210 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnu0202 
         Caption         =   "Productos"
      End
      Begin VB.Menu mnu0203 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu mnu0204 
         Caption         =   "Marcas"
      End
      Begin VB.Menu mnu0205 
         Caption         =   "Unidad de Medida"
      End
      Begin VB.Menu mnu0206 
         Caption         =   "Conceptos"
      End
      Begin VB.Menu mnu0207 
         Caption         =   "Niveles"
      End
      Begin VB.Menu mnu0208 
         Caption         =   "Almacenes"
      End
      Begin VB.Menu mnu0211 
         Caption         =   "Emp. Transporte"
      End
      Begin VB.Menu mnu0212 
         Caption         =   "Vehiculos"
      End
      Begin VB.Menu mnu0213 
         Caption         =   "Choferes"
      End
      Begin VB.Menu mnu0214 
         Caption         =   "Listas de Precios"
      End
      Begin VB.Menu mnu0215 
         Caption         =   "Formas de Pago"
      End
      Begin VB.Menu mnu0216 
         Caption         =   "Sucursales"
      End
      Begin VB.Menu mnu0217 
         Caption         =   "Cajas"
      End
      Begin VB.Menu mnu0218 
         Caption         =   "Tipos Mov. Caja"
      End
      Begin VB.Menu mnu0219 
         Caption         =   "Centros de Costo"
      End
      Begin VB.Menu mnu0220 
         Caption         =   "Almacenes para Ventas"
      End
      Begin VB.Menu mnu0221 
         Caption         =   "Grupos de Productos"
      End
      Begin VB.Menu mnu0222 
         Caption         =   "Numero Maximo Registros"
      End
      Begin VB.Menu mnu0223 
         Caption         =   "Zonas"
      End
      Begin VB.Menu mnu0224 
         Caption         =   "Ubigeo"
      End
      Begin VB.Menu mnu0225 
         Caption         =   "Motivos Traslados"
      End
      Begin VB.Menu mnu0226 
         Caption         =   "Tallas"
      End
      Begin VB.Menu mnu0227 
         Caption         =   "Tipos de Cambio"
      End
      Begin VB.Menu mnu0228 
         Caption         =   "Canales"
      End
      Begin VB.Menu mnu0229 
         Caption         =   "Datos"
      End
      Begin VB.Menu mnu0230 
         Caption         =   "Fórmulas"
      End
      Begin VB.Menu mnu0231 
         Caption         =   "Unidades de Producción"
      End
      Begin VB.Menu mnu0233 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnu0200 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu03 
      Caption         =   "R&eporte"
      Begin VB.Menu mnu0303 
         Caption         =   "Ventas"
         Begin VB.Menu mnu030301 
            Caption         =   "Ventas por Cliente"
         End
         Begin VB.Menu mnu030302 
            Caption         =   "Ventas por Producto"
         End
         Begin VB.Menu mnu030303 
            Caption         =   "Ventas por Responsable"
         End
         Begin VB.Menu mnu030305 
            Caption         =   "Ventas por Vendedor de Campo"
         End
         Begin VB.Menu mnu030306 
            Caption         =   "Ventas por Grupo de Productos"
         End
         Begin VB.Menu mnu030307 
            Caption         =   "Resumen de Ventas"
         End
         Begin VB.Menu mnu030308 
            Caption         =   "Ventas Gratuitas"
         End
         Begin VB.Menu mnu030309 
            Caption         =   "Ventas de Genética Líquida"
         End
         Begin VB.Menu mnu030310 
            Caption         =   "Ventas por Transferencias Gratuitas"
         End
         Begin VB.Menu mnu030399 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0308 
         Caption         =   "Registro de Ventas"
      End
      Begin VB.Menu mnu0312 
         Caption         =   "Stock - Ventas - Rotacion"
      End
      Begin VB.Menu mnu0313 
         Caption         =   "Caja"
         Begin VB.Menu mnu031301 
            Caption         =   "Liquidación de Caja - Detallado"
         End
         Begin VB.Menu mnu031302 
            Caption         =   "Documentos en Caja"
         End
         Begin VB.Menu mnu031399 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0314 
         Caption         =   "Compras"
         Begin VB.Menu mnu031401 
            Caption         =   "Registro de Movimientos - Compras"
         End
         Begin VB.Menu mnu031499 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0315 
         Caption         =   "Productos"
         Begin VB.Menu mnu031501 
            Caption         =   "Productos por Cliente"
         End
         Begin VB.Menu mnu031502 
            Caption         =   "Resumen Detallado por Producto"
         End
         Begin VB.Menu mnu031599 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0317 
         Caption         =   "Inventario"
         Begin VB.Menu mnu031701 
            Caption         =   "Kardex"
         End
         Begin VB.Menu mnu031702 
            Caption         =   "Kardex Por Talla"
         End
         Begin VB.Menu mnu031703 
            Caption         =   "Resumen de Stock"
         End
         Begin VB.Menu mnu031704 
            Caption         =   "Resumen de Stock Por Talla"
         End
         Begin VB.Menu mnu031705 
            Caption         =   "Estado de Documentos"
         End
         Begin VB.Menu mnu031706 
            Caption         =   "Consulta de Documentos"
         End
         Begin VB.Menu mnu031707 
            Caption         =   "Guías No Facturadas"
         End
         Begin VB.Menu mnu031708 
            Caption         =   "Despachos por Chofer"
         End
         Begin VB.Menu mnu031799 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0319 
         Caption         =   "Estadísticas"
         Begin VB.Menu mnu031901 
            Caption         =   "Estadístico de Ventas Mensual"
         End
         Begin VB.Menu mnu031902 
            Caption         =   "Estadístico de Ventas"
         End
         Begin VB.Menu mnu031903 
            Caption         =   "Estadístico de Ventas - Empresas Relacionadas"
         End
         Begin VB.Menu mnu031904 
            Caption         =   "Rentabilidad"
         End
         Begin VB.Menu mnu031999 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0318 
         Caption         =   "Ventas por Sucursal por Hora"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu0322 
         Caption         =   "Ventas por Cliente con dscto especial"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu0334 
         Caption         =   "Venta de Reproductores"
      End
      Begin VB.Menu mnu0337 
         Caption         =   "Comisiones"
      End
      Begin VB.Menu mnu0340 
         Caption         =   "Por Centro de Costo"
      End
      Begin VB.Menu mnu0344 
         Caption         =   "Atribuciones"
      End
      Begin VB.Menu mnu0345 
         Caption         =   "Ranking de Ventas"
      End
      Begin VB.Menu mnu0346 
         Caption         =   "Comisiones Por Vendedor de Campo"
      End
      Begin VB.Menu mnu0347 
         Caption         =   "Ranking de Ventas por Línea"
      End
      Begin VB.Menu mnu0348 
         Caption         =   "Situación del Pedido"
      End
      Begin VB.Menu mnu0349 
         Caption         =   "Situación de la Cotización"
      End
      Begin VB.Menu mnu0331 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu09 
      Caption         =   "&Seguridad"
      Begin VB.Menu mnu0901 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnu0902 
         Caption         =   "Perfiles"
      End
      Begin VB.Menu mnu0903 
         Caption         =   "Claves de Autorizacion"
      End
      Begin VB.Menu mnu0905 
         Caption         =   "Cierre del Mes"
      End
      Begin VB.Menu mnu0906 
         Caption         =   "Auditoría"
         Begin VB.Menu mnu090601 
            Caption         =   "Registro de Movimientos"
         End
         Begin VB.Menu mnu090699 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu0913 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu11 
      Caption         =   "&Utilitario"
      Begin VB.Menu mnu1101 
         Caption         =   "Series por Usuario"
      End
      Begin VB.Menu mnu1102 
         Caption         =   "Cambio Comision del Documento"
      End
      Begin VB.Menu mnu1103 
         Caption         =   "Procesar Tipos de Cambio"
      End
      Begin VB.Menu mnu1104 
         Caption         =   "Asientos Contables"
         Begin VB.Menu mnu110401 
            Caption         =   "Generar"
         End
         Begin VB.Menu mnu110402 
            Caption         =   "Consulta"
         End
         Begin VB.Menu mnu110403 
            Caption         =   "Consistencia"
            Begin VB.Menu mnu11040301 
               Caption         =   "Sin Cuadrar"
            End
            Begin VB.Menu mnu11040302 
               Caption         =   "De Cuentas"
            End
            Begin VB.Menu mnu11040399 
               Caption         =   "-"
            End
         End
         Begin VB.Menu mnu110404 
            Caption         =   "Transferir"
         End
         Begin VB.Menu mnu110405 
            Caption         =   "Eliminar"
         End
         Begin VB.Menu mnu110499 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu1105 
         Caption         =   "Cambio de Caja"
      End
      Begin VB.Menu mnu1106 
         Caption         =   "Liberar Documento"
      End
      Begin VB.Menu mnu1107 
         Caption         =   "Cambio de vendedor"
      End
      Begin VB.Menu mnu1108 
         Caption         =   "Transferir a Bancos"
      End
      Begin VB.Menu mnu1109 
         Caption         =   "Generar PDB"
      End
      Begin VB.Menu mnu1110 
         Caption         =   "Importar Productos"
      End
      Begin VB.Menu mnu1111 
         Caption         =   "Ventas Electrónicas"
         Begin VB.Menu mnu111101 
            Caption         =   "Comunicación de Baja"
         End
         Begin VB.Menu mnu111102 
            Caption         =   "Resumen Diario de Boletas"
         End
         Begin VB.Menu mnu111100 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnu1112 
         Caption         =   "Consulta de Entidades"
      End
      Begin VB.Menu mnu1113 
         Caption         =   "Consulta de Clientes"
      End
      Begin VB.Menu mnu1114 
         Caption         =   "Consulta de Proveedores"
      End
      Begin VB.Menu mnu1115 
         Caption         =   "Consulta de Productos"
      End
      Begin VB.Menu mnu1116 
         Caption         =   "Direcciones de Recojo"
      End
      Begin VB.Menu mnu1199 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu08 
      Caption         =   "Configuracion"
      Begin VB.Menu mnu0801 
         Caption         =   "Conf. Documentos de Ventas"
      End
      Begin VB.Menu mnu0804 
         Caption         =   "Conf. Impresion de Documentos de Venta"
      End
      Begin VB.Menu mnu0805 
         Caption         =   "Conf. Impresion de Motivos Guias"
      End
      Begin VB.Menu mnu0803 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnu0802 
         Caption         =   "Series por Usuario"
      End
      Begin VB.Menu mnu0806 
         Caption         =   "Variables del Sistema"
      End
      Begin VB.Menu mnu0807 
         Caption         =   "Tipos Niveles"
      End
      Begin VB.Menu mnu0808 
         Caption         =   "Tipos Mov. Caja"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu0809 
         Caption         =   "Documentos a Importar por Documento"
      End
      Begin VB.Menu mnu0810 
         Caption         =   "Conf. Impresion de Recibos - Caja"
      End
      Begin VB.Menu mnu0811 
         Caption         =   "Conf. Etiquetas a Imprimir por Doc."
      End
      Begin VB.Menu mnu0812 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu04 
      Caption         =   "V&entana"
      WindowList      =   -1  'True
      Begin VB.Menu mnu0401 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mnu0402 
         Caption         =   "Cerrar Todo"
      End
      Begin VB.Menu mnu0403 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu05 
      Caption         =   "&Ayuda"
      Visible         =   0   'False
      Begin VB.Menu mnu0501 
         Caption         =   "Acerca de"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu0502 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu06 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
Dim F As New frmDocVentas
Dim OC As New frmDocVentas_OC
Dim X As Object
Dim strTD As String

    Select Case pGroup.ObjectName
    Case "VENTAS"
        Select Case pLink.item.ObjectName
            Case "FACTURA"
                strTD = "01"
            Case "BOLETA"
                strTD = "03"
            Case "PEDIDO"
                strTD = "40"
            Case "COTIZACION"
                strTD = "92"
            Case "GUIA"
                strTD = "86"
            Case "CREDITO"
                strTD = "07"
            Case "DEBITO"
                strTD = "08"
            Case "TICKET"
                strTD = "12"
            Case "NOTAPEDIDO"
                strTD = "90"
            Case "NOTAVENTA"
                strTD = "56"
            Case "SEPARACION"
                strTD = "91"
            Case "BANDEJAPEDIDO"
                Dim B As New frmBandejaPedidos
                Load B
                B.Show
            Case "DESCUENTO"
                strTD = "89"
            Case "Guias Madres"
                Dim C As New FrmMantGuiasM
                Load C
                C.Show
            Case "Liquidaciones de Ventas"
                Dim D As New FrmLiquidacion
                Load D
                D.Show
        End Select
        If strTD <> "" Then
            F.strTipoDoc = strTD
            F.top = 0
            Load F
            F.Show
        End If
    Case "INVENTARIO"
        Select Case pLink.item.ObjectName
        Case "VINGRESO"
            Set X = New frmVale
            X.indVale = "I"
            Load X
            X.Show
        Case "VSALIDA"
            Set X = New frmVale
            X.indVale = "S"
            Load X
            X.Show
        Case "TRANSFERENCIA"
            Set X = New frmValeTrans
            Load X
            X.Show
        Case "PERIODO"
            Set X = New frmPeriodosINV
            Load X
            X.Show
        Case "COSTO"
            Set X = New frmProcesaCostosVales
            Load X
            X.Show
        Case "SALDOS"
            Set X = New FrmProcesaSaldos
            Load X
            X.Show
        Case "OCOMPRA"
            OC.strTipoDoc = "94"
            OC.top = 0
            Load OC
            OC.Show
        End Select
    End Select

End Sub

Private Sub MDIForm_Load()
On Error GoTo Err
Dim codPerfil As String
Dim StrMsgError As String
Dim C As Control
Dim rst         As New ADODB.Recordset
    '---- SI ES LLAMADO DESDE EL MENU PRINCIPAL
    If wingreso = False Then
        StrcodSistema = "01"
        Captura_parametros
        
        strRutaLogo = App.Path & "\Logo\LogoPW.jpg"
        App.Title = "SIAC"
        abrirConexion StrMsgError
        glsPersonaEmpresa = traerCampo("empresas", "idPersona", "idEmpresa", glsEmpresa, False)
    End If

    cargarImagenFondo
    gStrRutaRpts = App.Path + "\Reportes\"
    
    cargarParametrosSistema StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    glsCodPeriodoINV = traerCampo("periodosinv", "idPeriodoInv", "estPeriodoInv", "ACT", True, " idSucursal = '" & glsSucursal & "'")
    
    glsAlmVentas = traerCampo("almacenesvtas", "idAlmacen", "idSucursal", glsSucursal, True)
    
    If glsCodPeriodoINV = "" Then
        MsgBox "La Sucursal no tiene un periodo abierto", vbInformation, App.Title
    End If
    
    strArregloMes(1) = "ENERO"
    strArregloMes(2) = "FEBRERO"
    strArregloMes(3) = "MARZO"
    strArregloMes(4) = "ABRIL"
    strArregloMes(5) = "MAYO"
    strArregloMes(6) = "JUNIO"
    strArregloMes(7) = "JULIO"
    strArregloMes(8) = "AGOSTO"
    strArregloMes(9) = "SEPTIEMBRE"
    strArregloMes(10) = "OCTUBRE"
    strArregloMes(11) = "NOVIEMBRE"
    strArregloMes(12) = "DICIEMBRE"
    
    If indAdmin Then
        For Each C In frmPrincipal.Controls
            If left(C.Name, 3) = "mnu" Then
               If C.Enabled = False Then
                  C.Enabled = True
               End If
            End If
        Next
        
        StatusBar1.Panels(1).Text = "USUARIO: ADMINISTRADOR"
        StatusBar1.Panels(2).Text = "PERFIL: ADMINISTRADORES"
    
    Else
        validaMenu Me, glsUser, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
        If rst.State = 1 Then rst.Close: Set rst = Nothing
        rst.Open "EXEC spu_Usuario_Perfil '" & glsUser & "','" & glsEmpresa & "','" & glsSucursal & "','" & StrcodSistema & "' ", Cn, adOpenStatic, adLockOptimistic
        
        If Not rst.EOF Then
        
'            StatusBar1.Panels(1).Text = "USUARIO: " & traerCampo("personas", "GlsPersona", "idpersona", glsUser, False)
'            StatusBar1.Panels(2).Text = "PERFIL: " & traerCampo("Perfil A Inner Join PerfilesPorUsuario B On A.IdEmpresa = B.IdEmpresa And A.IdPerfil = B.IdPerfil And A.CodSistema = B.CodSistema", "A.GlsPerfil", "B.IdUsuario", glsUser, False, "A.IdEmpresa = '" & glsEmpresa & "' And B.CodSistema = '" & StrcodSistema & "'")
'            StatusBar1.Panels(3).Text = "SUCURSAL: " & traerCampo("personas", "GlsPersona", "idpersona", glsSucursal, False)
'            Me.Caption = Me.Caption & " - " & traerCampo("empresas", "GlsEmpresa", "idEmpresa", glsEmpresa, False)
            
            StatusBar1.Panels(1).Text = "USUARIO: " & Trim("" & rst.Fields("GlsPersona"))
            StatusBar1.Panels(2).Text = "PERFIL: " & Trim("" & rst.Fields("GlsPerfil"))
            StatusBar1.Panels(3).Text = "SUCURSAL: " & Trim("" & rst.Fields("glsSucursal"))
            Me.Caption = Me.Caption & " - " & Trim("" & rst.Fields("GlsEmpresa"))
            
        End If
    End If
    
    glsNumNiveles = traerCampo("tiposniveles", "Count(idTipoNivel)", "idEmpresa", glsEmpresa, False)
    StatusBar1.Panels(4).Text = "  FECHA: " & getFechaSistema
Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("¿Seguro de salir del sistema?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Cancel = 1
    Else
        
        For Each Obj In Forms
            If Obj.Name <> "frmPrincipal" Then
                Unload Obj
            End If
        Next
        
        Cn.Close
        Set Cn = Nothing
        End
    End If

End Sub


Private Sub mnu0229_Click()
Dim F As New frmMantDatos

    Load F
    F.Show

End Sub

Private Sub mnu0230_Click()
'Dim F As New frmMantCombos
'
'    Load F
'    F.Show

End Sub

Private Sub mnu0231_Click()
Dim F As New FrmMantUnidadProduccion
    
    Load F
    F.Show

End Sub

Private Sub mnu0233_Click()
Dim F As New FrmMantDocumentos

    Load F
    F.Show

End Sub

Private Sub mnu030301_Click()
Dim F As New FrmVentasxCliente

    Load F
    F.Show

End Sub

Private Sub mnu030302_Click()
Dim F As New FrmRptVentasPorProducto

    Load F
    F.Show
    
End Sub

Private Sub mnu030303_Click()
Dim F As New frmRptVentasxResponsable

    Load F
    F.Show
    
End Sub

Private Sub mnu030304_Click()
Dim F As New frmRptVentasporResponsableTDoc
    '--- 02/09
    '--- SE ESTA DESHABILITANDO ESTE REPORTE, ES EL MISMO
    '--- QUE VENTAS POR RESPONSABLE
    
    Load F
    F.Show
    
End Sub

Private Sub mnu030305_Click()
Dim F As New FrmRptVentasporVendedorCampo

    Load F
    F.Show
    
End Sub

Private Sub mnu030306_Click()
Dim F As New frmRptVentasporGrupoProductos

    Load F
    F.Show
    
End Sub

Private Sub mnu030307_Click()
Dim F As New FrmResumenVentas

    Load F
    F.Show
    
End Sub

Private Sub mnu030308_Click()
Dim F As New FrmRptVtaGratuitas
    
    Load F
    F.Show
    
End Sub

Private Sub mnu030309_Click()
Dim F As New frmRptVentaGeneticaLiquida

    Load F
    F.Show

End Sub

Private Sub mnu030310_Click()
Dim F As New FrmRptTransGratuitas
    
    Load F
    F.Show

End Sub

Private Sub mnu031704_Click()
Dim F As New FrmRptInventarioValorizado_Lotes
    
    Load F
    F.Show
End Sub

Private Sub mnu031705_Click()
Dim F As New frmRptEstadoDocumentos

    Load F
    F.Show

End Sub

Private Sub mnu031904_Click()
Dim F As New FrmRptRentabilidad

    Load F
    F.Show
    
End Sub

Private Sub mnu0334_Click()
Dim F As New FrmrptReproductores

    Load F
    F.Show

End Sub

Private Sub mnu031902_Click()
Dim F As New FrmEstadisticos
        
    sw_estadistico = False  '--- Estadístico de Ventas para TODOS los clientes
    Load F
    F.Show

End Sub

Private Sub mnu0337_Click()
Dim F As New FrmRptComision

    Load F
    F.Show

End Sub

Private Sub mnu031903_Click()
Dim F As New FrmEstadisticos
        
    sw_estadistico = True  '--- Estadístico de Ventas - Empresas Relacionadas
    Load F
    F.Show
    
End Sub

Private Sub mnu0340_Click()
Dim F As New FrmRptPorCentroCosto

    Load F
    F.Show

End Sub

Private Sub mnu031706_Click()
Dim F As New FrmSecuenciaDocumentos

    Load F
    F.Show

End Sub

Private Sub mnu031707_Click()
Dim F As New FrmGuiasNoFacturadas

    Load F
    F.Show

End Sub

Private Sub mnu0344_Click()
Dim F As New FrmRptAtribuciones

    Load F
    F.Show

End Sub

Private Sub mnu0345_Click()
Dim F As New frmRankingVentas

    Load F
    F.Show
    
End Sub

Private Sub mnu0346_Click()
Dim F As New FrmRptVentasporVendedorCampo_Comisión

    Load F
    F.Show
    
End Sub

Private Sub mnu0347_Click()
Dim F As New FrmRankingVentasPorLinea

    Load F
    F.Show
    
End Sub

Private Sub mnu0348_Click()
Dim F As New frmRptSituacionPedido

    Load F
    F.Show
    
End Sub

Private Sub mnu0349_Click()
Dim F As New FrmRptSituacionCotizacion

    Load F
    F.Show

End Sub

Private Sub mnu0401_Click()
    
    Arrange vbCascade

End Sub

Private Sub mnu0402_Click()
Dim Obj As Object

    For Each Obj In Forms
        If Obj.Name <> "frmPrincipal" Then
            Unload Obj
        End If
    Next

End Sub

Private Sub mnu0501_Click()
    
    frmAbout.Show 1

End Sub

Private Sub mnu06_Click()
    
    Unload Me

End Sub

Private Sub mnu0715_Click()
Dim F As New FrmMantGuiasM

    Load F
    F.Show

End Sub

Private Sub mnu0716_Click()
Dim F As New FrmLiquidacion

    Load F
    F.Show

End Sub

Private Sub mnu0717_Click()
Dim F As New FrmAtribucion
Dim F1 As New frmDocVentas

    If leeParametro("ATRIBUCIONES_VENTAS") = "S" Then
        F1.strTipoDoc = "25"
        F1.IndAtribucionNC = 0
        F1.top = 0
        Load F1
        F1.Show
    Else
        Load F
        F.Show
    End If
    
End Sub

Private Sub mnu0718_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "25"
    F.IndAtribucionNC = 1
    F.top = 0
    Load F
    F.Show

End Sub

Private Sub mnu0807_Click()

    Load frmMantTiposNiveles
    frmMantTiposNiveles.Show

End Sub

Private Sub mnu0808_Click()
    
    frmConfTipoMovCaja.Show

End Sub

Private Sub mnu0809_Click()
    
    frmDocumentoExportar.Show

End Sub

Private Sub mnu0810_Click()
    
    frmConfImpRecibos.Show

End Sub

Private Sub mnu0811_Click()
    
    frmConfObjEtiquetasDoc.Show

End Sub

Private Sub mnu0812_Click()
    
    frmSeriesDocumento.Show 1

End Sub

Private Sub mnu0201_Click()
Dim F As New frmMantPersona

    Load F
    F.Show

End Sub

Private Sub mnu0202_Click()
Dim F As New frmMantProductos

    Load F
    F.Show

End Sub

Private Sub mnu0203_Click()
Dim F As New frmMantProveedores

    Load F
    F.Show

End Sub

Private Sub mnu0204_Click()
Dim F As New frmMantMarcas

    Load F
    F.Show

End Sub

Private Sub mnu0205_Click()
Dim F As New frmMantUM

    Load F
    F.Show

End Sub

Private Sub mnu0206_Click()
Dim F As New frmMantConceptos

    Load F
    F.Show

End Sub

Private Sub mnu0207_Click()
Dim F As New frmMantNiveles

    Load F
    F.Show

End Sub

Private Sub mnu0208_Click()
Dim F As New frmMantAlmacenes

    Load F
    F.Show

End Sub

Private Sub mnu0209_Click()
Dim F As New frmMantClientes

    Load F
    F.Show

End Sub

Private Sub mnu0210_Click()
Dim F As New frmMantVendedores

    Load F
    F.Show

End Sub

Private Sub mnu0211_Click()
Dim F As New frmMantEmpTrans

    Load F
    F.Show

End Sub

Private Sub mnu0212_Click()
Dim F As New frmMantVehiculos

    Load F
    F.Show

End Sub

Private Sub mnu0213_Click()
Dim F As New frmMantChoferes

    Load F
    F.Show

End Sub

Private Sub mnu0214_Click()
Dim F As New frmMantListaPrecios

    Load F
    F.Show

End Sub

Private Sub mnu0215_Click()
Dim F As New frmMantFormasPago

    Load F
    F.Show

End Sub

Private Sub mnu0216_Click()
Dim F As New frmMantSucursales

    Load F
    F.Show

End Sub

Private Sub mnu0217_Click()
Dim F As New frmMantCajas

    Load F
    F.Show

End Sub

Private Sub mnu0218_Click()
Dim F As New frmConfTipoMovCaja

    Load F
    F.Show

End Sub

Private Sub mnu0219_Click()
Dim F As New frmMantCentroCosto

    Load F
    F.Show

End Sub

Private Sub mnu0220_Click()
Dim F As New frmMantAlmacenesVtas

    Load F
    F.Show

End Sub

Private Sub mnu0221_Click()
Dim F As New frmMantGruposProductos

    Load F
    F.Show

End Sub

Private Sub mnu0222_Click()
Dim F As New frmSeriesDocumento

    Load F
    F.Show

End Sub

Private Sub mnu0223_Click()
Dim F As New frmMantZonas

    Load F
    F.Show

End Sub

Private Sub mnu0224_Click()
Dim F As New frmMantUbigeo

    Load F
    F.Show

End Sub

Private Sub mnu0225_Click()
Dim F As New frmMantMotivosGuia

    Load F
    F.Show

End Sub

Private Sub mnu0226_Click()
Dim F As New FrmMantLotes 'frmMantTallasPesos

    Load F
    F.Show

End Sub

Private Sub mnu0227_Click()
Dim F As New frmMantTiposCambio

    Load F
    F.Show

End Sub

Private Sub mnu0228_Click()
Dim F As New FrmMant_Canal
    
    Load F
    F.Show

End Sub

Private Sub mnu0308_Click()
Dim F As New frmRptRegVentas

    Load F
    F.Show

End Sub

'Private Sub mnu0312_Click()
'Dim F As New frmReporteStockVentasRotacion
'
'    Load F
'    F.Show
'
'End Sub

Private Sub mnu031301_Click()
Dim F As New frmRptLiquidacionCajaDetallado

    Load F
    F.Show
    
End Sub

Private Sub mnu031701_Click()
Dim F As New FrmKardex
    
    F.Show
    
End Sub

Private Sub mnu0318_Click()
Dim F As New frmReportes

    F.GlsForm = "Ventas por Sucursal por Hora"
    F.NumerosFrames = "1,5"
    F.GlsReporte = "rptVentasSucursalHoras.rpt"
    Load F
    F.Show

End Sub

Private Sub mnu031702_Click()
Dim F As New FrmKardexPorLote
    
    Load F
    F.Show

End Sub

Private Sub mnu0322_Click()
Dim F As New frmReportes

    F.GlsForm = "Ventas por Cliente con dscto especial"
    F.NumerosFrames = "1,4,5"
    F.GlsReporte = "rptVentasPorClienteDctoEspecial.rpt"
    Load F
    F.Show

End Sub

Private Sub mnu031708_Click()
Dim F As New FrmDespachos

    Load F
    F.Show

End Sub

Private Sub mnu031901_Click()
Dim F As New FrmestadisticoVentasMensual

    Load F
    F.Show

End Sub

Private Sub mnu031401_Click()
Dim F As New frmRptCompras

    Load F
    F.Show

End Sub

Private Sub mnu031402_Click()
Dim F As New FrmClientesPorNivel

    Load F
    F.Show

End Sub

Private Sub mnu031302_Click()
Dim F As New frmrptcaja

    Load F
    F.Show

End Sub

Private Sub mnu031703_Click()
Dim F As New FrmRptInventarioValorizado
    
    Load F
    F.Show
End Sub

Private Sub mnu031501_Click()
Dim F As New FrmProductosPorClientes

    Load F
    F.Show

End Sub

Private Sub mnu031502_Click()
Dim F As New FrmDetalladoPorProducto

    Load F
    F.Show

End Sub

Private Sub mnu0701_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "01"
    F.top = 0
    Load F
    F.Show

End Sub

Private Sub mnu0702_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "03"
    Load F
    F.Show

End Sub

Private Sub mnu0703_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "40"
    Load F
    F.Show

End Sub

Private Sub mnu0705_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "86"
    Load F
    F.Show

End Sub

Private Sub mnu0706_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "07"
    Load F
    F.Show

End Sub

Private Sub mnu0801_Click()
    
    frmConfDocVentas.Show

End Sub

Private Sub mnu0802_Click()
Dim F As New frmSeriesUsuario

    Load F
    F.Show

End Sub

Private Sub mnu0803_Click()
Dim F As New frmConfDocumentos

    Load F
    F.Show

End Sub

Private Sub mnu0804_Click()
    
    frmConfiguraImpresion.Show

End Sub

Private Sub mnu0805_Click()

    frmMotivosGuiaImp.Show

End Sub

Private Sub mnu0806_Click()
    
    frmVariablesSistema.Show

End Sub

Private Sub mnu0708_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "12"
    Load F
    F.Show

End Sub

Private Sub mnu0709_Click()
Dim F As New frmBandejaPedidos

    Load F
    F.Show

End Sub

Private Sub mnu0710_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "56"
    Load F
    F.Show

End Sub

Private Sub mnu0711_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "91"
    Load F
    F.Show

End Sub

Private Sub mnu0712_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "92"
    Load F
    F.Show

End Sub

Private Sub mnu0713_Click()
Dim F As New frmDocVentas

    F.strTipoDoc = "89"
    Load F
    F.Show

End Sub

Private Sub mnu0901_Click()
Dim F As New frmMantUsuarios

    Load F
    F.Show

End Sub

Private Sub mnu0902_Click()
Dim F As New frmMantPerfiles

    Load F
    F.Show

End Sub

Private Sub mnu0903_Click()
Dim F As New frmMantClaves

    Load F
    F.Show

End Sub

Private Sub mnu0905_Click()
Dim F As New frmCierreMes

    Load F
    F.Show

End Sub

Private Sub mnu090601_Click()
Dim F As New frmAuditoriaRegMovVentas

    Load F
    F.Show

End Sub

Private Sub mnu1001_Click()
Dim F               As Object

    Set F = New frmVale
    F.indVale = "I"
    Load F
    F.Show

End Sub

Private Sub mnu1002_Click()
Dim F               As Object

    Set F = New frmVale

    F.indVale = "S"
    Load F
    F.Show
End Sub

Private Sub mnu1003_Click()
Dim F As New frmValeTrans
    
    Load F
    F.Show

End Sub

Private Sub mnu1004_Click()
Dim F As New frmPeriodosINV
    
    Load F
    F.Show

End Sub

Private Sub mnu1005_Click()
Dim F As New frmProcesaCostosVales
    
    Load F
    F.Show
End Sub

Private Sub mnu1006_Click()
Dim F As New FrmProcesaSaldos

    Load F
    F.Show

End Sub

Private Sub cargarImagenFondo()
Dim strRutaLogo As String
Dim strImagen As String
Dim imagen As IPictureDisp

    strImagen = traerCampo("parametros", "ValParametro", "glsParametro", "IMAGEN_FONDO", True)
    strRutaLogo = App.Path & "\Logo\" & strImagen
    Set imagen = LoadPicture(strRutaLogo)
    Me.Picture = imagen

End Sub

Private Sub mnu1007_Click()
Dim F As New frmDocVentas_OC

    F.strTipoDoc = "94"
    F.top = 0
    Load F
    F.Show
    
End Sub

Private Sub mnu1101_Click()
Dim F As New frmSeriesUsuario

    Load F
    F.Show

End Sub

Private Sub mnu1102_Click()
Dim F As New frmCambiarComisionesDoc

    Load F
    F.Show

End Sub

Private Sub mnu1103_Click()
Dim F As New FrmProcesarTC

    Load F
    F.Show

End Sub

Private Sub mnu110401_Click()
Dim F As New frmAsientoContableGenerar

    Load F
    F.Show

End Sub

Private Sub mnu110402_Click()
Dim F As New frmAsientoContableConsultar

    Load F
    F.Show

End Sub

Private Sub mnu11040301_Click()
Dim F As New frmAsientoContableConsistenciaSinCuadrar

    Load F
    F.Show

End Sub

Private Sub mnu11040302_Click()
Dim F As New frmAsientoContableConsistenciaDeCuentas

    Load F
    F.Show

End Sub

Private Sub mnu110404_Click()
Dim F As New frmAsientoContableTransferir

    Load F
    F.Show

End Sub

Private Sub mnu110405_Click()
Dim F As New frmAsientoContableEliminar
    
    Load F
    F.Show

End Sub

Private Sub mnu1105_Click()
Dim F As New FrmCambioCaja

    Load F
    F.Show

End Sub

Private Sub mnu1106_Click()
Dim F As New FrmModificarEstado

    Load F
    F.Show

End Sub

Private Sub mnu1107_Click()
Dim F As New FrmCambioVendedorCampo
    
    Load F
    F.Show

End Sub

Private Sub mnu1108_Click()
Dim F As New FrmTransfiereBancos
    
    Load F
    F.Show

End Sub

Sub Captura_parametros()
Dim xContador As Integer, i As Integer
Dim xrecepciona As String
Dim xParametro1 As String
Dim xParametro2 As String
Dim xParametro3 As String
Dim cadena As String
    
    xrecepciona = Command
    cadena = ""
    xParametro1 = ""
    xParametro2 = ""
    xParametro3 = ""
    xContador = 0
    For i = 1 To Len(xrecepciona)
        If Mid(xrecepciona, i, 1) = "," Then
            xContador = xContador + 1
            Select Case xContador
                Case 1: xParametro1 = cadena
                Case 2: xParametro2 = cadena
                Case 3: xParametro3 = cadena
            End Select
            cadena = ""
        Else
            cadena = cadena & Mid(xrecepciona, i, 1)
        End If
    Next
    glsEmpresa = xParametro1
    glsSucursal = xParametro2
    glsUser = xParametro3
    
End Sub

Private Sub mnu1109_Click()
Dim F As New FrmGeneraPDB
    
    Load F
    F.Show

End Sub

Private Sub mnu1110_Click()
Dim F As New FrmImportaProductos
    
    Load F
    F.Show

End Sub

Private Sub mnu111101_Click()
Dim F As New FrmComunicacionBaja
    
    Load F
    F.Show

End Sub

Private Sub mnu111102_Click()
Dim F As New FrmResumenDiarioBoletas
    
    Load F
    F.Show

End Sub

Private Sub mnu1112_Click()
Dim F As New frmMantPersona_Consul
    
    Load F
    F.Show
End Sub

Private Sub mnu1113_Click()
Dim F As New frmMantClientes_Consul
    
    Load F
    F.Show
End Sub

Private Sub mnu1114_Click()
Dim F As New frmMantProveedores_Consul
    
    Load F
    F.Show
End Sub

Private Sub mnu1115_Click()
Dim F As New frmMantProductos_Consul
    
    Load F
    F.Show
End Sub

Private Sub mnu1116_Click()
Dim F As New frmMantDirRecojo
    
    Load F
    F.Show
End Sub
