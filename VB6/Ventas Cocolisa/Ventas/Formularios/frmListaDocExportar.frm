VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmListaDocExportar 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Documento"
   ClientHeight    =   8865
   ClientLeft      =   4560
   ClientTop       =   1200
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   750
      Top             =   4500
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
            Picture         =   "frmListaDocExportar.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   1164
      ButtonWidth     =   3175
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "            Aceptar            "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12510
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   12315
         Begin VB.CommandButton cmbAyudaCliente 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7110
            Picture         =   "frmListaDocExportar.frx":3518
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   230
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaDocumento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7110
            Picture         =   "frmListaDocExportar.frx":38A2
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   615
            Width           =   390
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   600
            TabIndex        =   4
            Top             =   1080
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmListaDocExportar.frx":3C2C
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   945
            TabIndex        =   7
            Top             =   620
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   8
            Container       =   "frmListaDocExportar.frx":3C48
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Top             =   620
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmListaDocExportar.frx":3C64
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Cliente 
            Height          =   315
            Left            =   945
            TabIndex        =   11
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   8
            Container       =   "frmListaDocExportar.frx":3C80
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Cliente 
            Height          =   315
            Left            =   1920
            TabIndex        =   12
            Top             =   240
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            BackColor       =   16777152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmListaDocExportar.frx":3C9C
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Desde 
            Height          =   315
            Left            =   8670
            TabIndex        =   14
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   132775937
            CurrentDate     =   38955
         End
         Begin MSComCtl2.DTPicker dtp_Hasta 
            Height          =   315
            Left            =   10860
            TabIndex        =   15
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   132775937
            CurrentDate     =   38955
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   8055
            TabIndex        =   17
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   10305
            TabIndex        =   16
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lbl_Cliente 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   75
            TabIndex        =   13
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   75
            TabIndex        =   9
            Top             =   660
            Width           =   810
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3825
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportar.frx":3CB8
         TabIndex        =   1
         Top             =   1485
         Width           =   12315
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   2535
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportar.frx":DCC1
         TabIndex        =   2
         Top             =   5475
         Width           =   12315
      End
   End
End
Attribute VB_Name = "frmListaDocExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsg             As New ADODB.Recordset
Private RsD             As New ADODB.Recordset
Private strTDExportar   As String
Dim indNuevoDoc         As Boolean

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaDocumento_Click()
    
    mostrarAyuda "DOCUMENOSEXP", txtCod_Documento, txtGls_Documento, " AND c.idDocumento = '" & strTDExportar & "'"

End Sub

Private Sub dtp_Desde_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub dtp_Hasta_Change()
On Error GoTo Err
Dim StrMsgError As String
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    strRptNum = ""
    strRptSerie = ""
    
    ConfGrid gLista, True, False, False, False
    ConfGrid gListaDetalle, True, False, False, False
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub listaDocVentas(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim strFiltroAprob As String

    '--- FORMATO GRID
    Set gLista.DataSource = Nothing
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    If RsD.State = 1 Then RsD.Close: Set RsD = Nothing
    
    '--- Cabecera
    rsg.Fields.Append "Item", adChar, 14, adFldRowID
    rsg.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsg.Fields.Append "idDocVentas", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    rsg.Fields.Append "idPerVendedor", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsVendedor", adVarChar, 185, adFldIsNullable
    rsg.Fields.Append "FecEmision", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "estDocVentas", adChar, 3, adFldIsNullable
    rsg.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    rsg.Fields.Append "TotalValorVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalIGVVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "TotalPrecioVenta", adDouble, 14, adFldIsNullable
    rsg.Fields.Append "numOrdenCompra", adVarChar, 200, adFldIsNullable
    rsg.Fields.Append "llegada", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "FecIniTraslado", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idFormaPago", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "TC", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idPerChofer", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "glsChofer", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "idPerEmpTrans", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsEmpTrans", adVarChar, 255, adFldIsNullable
    rsg.Fields.Append "idVehiculo", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "GlsVehiculo", adVarChar, 100, adFldIsNullable
    rsg.Fields.Append "Placa", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Marca", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Color", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Modelo", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "CodInsCrip", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "Brevete", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "rucEmpTrans", adVarChar, 30, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdLista", adVarChar, 8, adFldIsNullable
    rsg.Fields.Append "IdAlmacen", adVarChar, 8, adFldIsNullable
    rsg.Open , , adOpenKeyset, adLockOptimistic
    
    '--- Detalle
    RsD.Fields.Append "Item", adVarChar, 20, adFldRowID
    RsD.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    RsD.Fields.Append "idProducto", adChar, 8, adFldIsNullable
    RsD.Fields.Append "idCodFabricante", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "GlsProducto", adVarChar, 500, adFldIsNullable
    RsD.Fields.Append "idMarca", adChar, 8, adFldIsNullable
    RsD.Fields.Append "GlsMarca", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "idUM", adChar, 8, adFldIsNullable
    RsD.Fields.Append "GlsUM", adVarChar, 185, adFldIsNullable
    RsD.Fields.Append "Factor", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "Afecto", adInteger, 4, adFldIsNullable
    RsD.Fields.Append "Cantidad", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "Cantidad2", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "VVUnit", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "IGVUnit", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "PVUnit", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "TotalVVBruto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "TotalPVBruto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "PorDcto", adVarChar, 20, adFldIsNullable
    RsD.Fields.Append "DctoVV", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "DctoPV", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "TotalVVNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "TotalIGVNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "TotalPVNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "idTipoProducto", adChar, 5, adFldIsNullable
    RsD.Fields.Append "idMoneda", adChar, 3, adFldIsNullable
    RsD.Fields.Append "idDocVentas", adChar, 8, adFldIsNullable
    RsD.Fields.Append "idSerie", adChar, 4, adFldIsNullable
    RsD.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "CodigoRapido", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "idTallaPeso", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "ItemPro", adInteger, , adFldRowID
    RsD.Fields.Append "IvapUnit", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "TotalIvapNeto", adDouble, 14, adFldIsNullable
    RsD.Open , , adOpenKeyset, adLockOptimistic
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsVendedor LIKE '%" & strCond & "%'"
    End If
    strFiltroAprob = ""
        
    If traerCampo("Documentos", "Aprobacion", "iddocumento", txtCod_Documento.Text, False) = "S" Then
        strFiltroAprob = " AND indAprobado = '1' "
    Else
        strFiltroAprob = ""
    End If
    
    If txtCod_Documento.Text = "40" And traerCampo("Parametros", "ValParametro", "GlsParametro", "APRUEBA_PEDIDO_AUTOMATICO", True) = "1" Then
        
        csql = "Select ConCat(A.IdDocumento,A.IdDocVentas,A.IdSerie) As Item,A.IdFormaPago,A.IdDocVentas,A.IdSerie,A.IdPerVendedor," & _
                "A.GlsVendedor,A.FecEmision,A.EstDocVentas,A.IdMoneda,A.TotalValorVenta,A.TotalIGVVenta,A.TotalPrecioVenta,A.NumOrdenCompra," & _
                "A.Llegada,A.IdMoneda,A.FecIniTraslado,A.TipoCambio,A.IdPerChofer,A.GlsChofer,A.IdPerEmpTrans,A.GlsEmpTrans,A.IdVehiculo," & _
                "A.GlsVehiculo,A.Placa,A.Marca,A.Color,A.Modelo,A.CodInsCrip,A.Brevete,A.RucEmpTrans,A.IdCentroCosto,A.IdLista,A.IdAlmacen " & _
                "From Docventas A " & _
                "Inner Join CentrosCosto C " & _
                    "On A.IdEmpresa = C.IdEmpresa And A.IdCentroCosto = C.IdCentroCosto " & _
                "Left Join ValesConver B " & _
                    "On C.IdEmpresa = B.IdEmpresa And C.IdCentroCosto = B.IdCentroCosto " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' " & _
                "And A.IdDocumento = '" & txtCod_Documento.Text & "' And A.EstDocventas <> 'ANU' " & _
                "And A.IdPerCliente = '" & txtCod_Cliente.Text & "' " & _
                "And A.FecEmision BetWeen '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' And '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' " & _
                "And A.EstDocImportado <> 'S' And (C.IndGenerado = 'M' Or (B.IndCierre = '1' And B.EstValeConver <> 'ANU'))" & strFiltroAprob
    
    ElseIf txtCod_Documento.Text = "40" And leeParametro("VALIDA_IMPORTE_PEDIDO") = "1" Then
        
        If traerCampo("Documentos", "Aprobacion", "iddocumento", txtCod_Documento.Text, False) = "S" Then
            strFiltroAprob = " AND A.indAprobado = '1' "
        Else
            strFiltroAprob = ""
        End If
    
        csql = "Select ConCat(A.IdDocumento,A.IdDocVentas,A.IdSerie) As Item,A.IdFormaPago,A.IdDocVentas,A.IdSerie,A.IdPerVendedor," & _
                "A.GlsVendedor,A.FecEmision,A.EstDocVentas,A.IdMoneda,A.TotalValorVenta,A.TotalIGVVenta,A.TotalPrecioVenta,A.NumOrdenCompra," & _
                "A.Llegada,A.IdMoneda,A.FecIniTraslado,A.TipoCambio,A.IdPerChofer,A.GlsChofer,A.IdPerEmpTrans,A.GlsEmpTrans,A.IdVehiculo," & _
                "A.GlsVehiculo,A.Placa,A.Marca,A.Color,A.Modelo,A.CodInsCrip,A.Brevete,A.RucEmpTrans,A.IdCentroCosto,A.IdLista,A.TotalPrecioVenta TotalPrecioVentaOri,A.IdAlmacen " & _
                "From DocVentas A " & _
                "Left Join DocReferencia B " & _
                    "On A.IdEmpresa = B.IdEmpresa And A.IdDocumento = B.TipoDocReferencia And A.IdSerie = B.SerieDocReferencia " & _
                    "And A.IdDocVentas = B.NumDocReferencia " & _
                "Left Join DocVentas C " & _
                    "On B.IdEmpresa = C.IdEmpresa And B.TipoDocOrigen = C.IdDocumento And B.SerieDocOrigen = C.IdSerie And B.NumDocOrigen = C.IdDocVentas " & _
                    "And 'ANU' <> C.EstDocVentas " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & txtCod_Documento.Text & "' " & _
                "And A.EstDocventas <> 'ANU' And A.IdPerCliente = '" & txtCod_Cliente.Text & "' " & _
                "And A.FecEmision BetWeen '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' And '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' " & _
                strFiltroAprob & _
                "Group By A.IdEmpresa,A.IdDocumento,A.IdSerie,A.IdDocVentas " & _
                "Having TotalPrecioVentaOri > Sum(IfNull(C.TotalPrecioVenta,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1))"
                
    Else
        
        If glsValidaStock Then
            
            csql = "Call Spu_ListaDocExportar('" & glsEmpresa & "','','1','" & txtCod_Cliente.Text & "','" & txtCod_Documento.Text & "','','','','" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "','" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "','C')"
            
        Else
        
            csql = "Select ConCat(IdDocumento,IdDocVentas,IdSerie) As Item,IdFormaPago,IdDocVentas,IdSerie,IdPerVendedor,GlsVendedor,FecEmision," & _
                    "EstDocVentas,IdMoneda,TotalValorVenta,TotalIGVVenta,TotalPrecioVenta,NumOrdenCompra,Llegada,IdMoneda,FecIniTraslado,TipoCambio," & _
                    "IdPerChofer,GlsChofer,IdPerEmpTrans,GlsEmpTrans,IdVehiculo,GlsVehiculo,Placa,Marca,Color,Modelo,CodInsCrip,Brevete,RucEmpTrans," & _
                    "IdCentroCosto,IdLista,IdAlmacen " & _
                    "From Docventas " & _
                    "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And IdDocumento = '" & txtCod_Documento.Text & "' " & _
                    "And EstDocventas <> 'ANU' And IdPerCliente = '" & txtCod_Cliente.Text & "' " & _
                    "And CAST(FecEmision AS DATE) BetWeen CAST('" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' AS DATE) And CAST('" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' AS DATE) " & _
                    "And EstDocImportado <> 'S'" & strFiltroAprob
                
        End If
        
                
    End If
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF And Not rst.BOF Then
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("chkMarca") = 0
            rsg.Fields("idDocVentas") = rst.Fields("idDocVentas")
            rsg.Fields("idSerie") = rst.Fields("idSerie")
            rsg.Fields("idPerVendedor") = rst.Fields("idPerVendedor")
            rsg.Fields("GlsVendedor") = rst.Fields("GlsVendedor")
            rsg.Fields("FecEmision") = Format(rst.Fields("FecEmision"), "dd/mm/yyyy")
            rsg.Fields("estDocVentas") = rst.Fields("estDocVentas")
            rsg.Fields("idMoneda") = rst.Fields("idMoneda")
            rsg.Fields("TotalValorVenta") = rst.Fields("TotalValorVenta")
            rsg.Fields("TotalIGVVenta") = rst.Fields("TotalIGVVenta")
            rsg.Fields("TotalPrecioVenta") = rst.Fields("TotalPrecioVenta")
            rsg.Fields("numOrdenCompra") = rst.Fields("numOrdenCompra")
            rsg.Fields("llegada") = rst.Fields("llegada")
            rsg.Fields("idmoneda") = rst.Fields("idmoneda")
            rsg.Fields("FecIniTraslado") = Format(rst.Fields("FecIniTraslado"), "dd/mm/yyyy")
            rsg.Fields("idFormaPago") = Trim("" & rst.Fields("idFormaPago"))
            rsg.Fields("TC") = Trim("" & rst.Fields("TipoCambio"))
            rsg.Fields("idPerChofer") = Trim("" & rst.Fields("idPerChofer"))
            rsg.Fields("glsChofer") = Trim("" & rst.Fields("glsChofer"))
            rsg.Fields("idPerEmpTrans") = Trim("" & rst.Fields("idPerEmpTrans"))
            rsg.Fields("GlsEmpTrans") = Trim("" & rst.Fields("GlsEmpTrans"))
            rsg.Fields("idVehiculo") = Trim("" & rst.Fields("idVehiculo"))
            rsg.Fields("GlsVehiculo") = Trim("" & rst.Fields("GlsVehiculo"))
            rsg.Fields("Placa") = Trim("" & rst.Fields("Placa"))
            rsg.Fields("Marca") = Trim("" & rst.Fields("Marca"))
            rsg.Fields("Color") = Trim("" & rst.Fields("Color"))
            rsg.Fields("Modelo") = Trim("" & rst.Fields("Modelo"))
            rsg.Fields("CodInsCrip") = Trim("" & rst.Fields("CodInsCrip"))
            rsg.Fields("Brevete") = Trim("" & rst.Fields("Brevete"))
            rsg.Fields("rucEmpTrans") = Trim("" & rst.Fields("rucEmpTrans"))
            rsg.Fields("IdCentroCosto") = Trim("" & rst.Fields("IdCentroCosto"))
            rsg.Fields("IdLista") = Trim("" & rst.Fields("IdLista"))
            rsg.Fields("IdAlmacen") = Trim("" & rst.Fields("IdAlmacen"))
            rst.MoveNext
        Loop

        mostrarDatosGridSQL gLista, rsg, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Me.Refresh
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
    Resume
End Sub

Private Sub ListaDetalle(ByRef StrMsgError As String)
Dim rst                                     As New ADODB.Recordset
Dim strCond                                 As String
Dim indExisteDoc                            As Boolean
Dim strNumDoc                               As String
Dim strSerie                                As String
Dim NValoresPedido()                    As Double
On Error GoTo Err

    '--- FORMATO GRID
    strNumDoc = gLista.Columns.ColumnByFieldName("idDocVentas").Value
    strSerie = gLista.Columns.ColumnByFieldName("idSerie").Value
    
    '--- Validamos si ya adicionamos el detalle
    gListaDetalle.Dataset.Filter = ""
    gListaDetalle.Dataset.Filtered = True
    indExisteDoc = False
    
    Set gListaDetalle.DataSource = Nothing
    gListaDetalle.Dataset.DisableControls
    
    If RsD.RecordCount > 0 Then RsD.MoveFirst
    Do While Not RsD.EOF
        If RsD.Fields("idDocVentas") = strNumDoc And RsD.Fields("idSerie") = strSerie Then
            indExisteDoc = True
            Exit Do
        End If
        RsD.MoveNext
    Loop
    
    If indExisteDoc = False Then
        
        If leeParametro("VALIDA_IMPORTE_PEDIDO") = "1" Then
        
            csql = "Select A.Item,A.IdProducto,A.IdCodFabricante,A.GlsProducto,A.IdMarca,A.GlsMarca,A.IdUM,A.GlsUM,A.Factor,A.Afecto," & _
                    "A.Cantidad,A.VVUnit - Sum(IfNull(D.VVUnit,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) VVUnit,A.IGVUnit - Sum(IfNull(D.IGVUnit,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) IGVUnit," & _
                    "A.PVUnit - Sum(IfNull(D.PVUnit,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) PVUnit,A.TotalVVBruto - Sum(IfNull(D.TotalVVBruto,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) TotalVVBruto," & _
                    "A.TotalPVBruto - Sum(IfNull(D.TotalPVBruto,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) TotalPVBruto,A.PorDcto,A.DctoVV,A.DctoPV," & _
                    "A.TotalVVNeto - Sum(IfNull(D.TotalVVNeto,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) TotalVVNeto," & _
                    "A.TotalIGVNeto - Sum(IfNull(D.TotalIGVNeto,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) TotalIGVNeto," & _
                    "A.TotalPVNeto - Sum(IfNull(D.TotalPVNeto,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1)) TotalPVNeto,A.IdTipoProducto,A.IdMoneda," & _
                    "A.NumLote,A.FecVencProd,A.VVUnitLista,A.PVUnitLista,A.VVUnitNeto,A.PVUnitNeto,A.Cantidad2,A.CodigoRapido,A.IdTallaPeso,A.ItemPro," & _
                    "A.TotalPVBruto TotalPVBrutoOri,A.IvapUnit,A.TotalIvapNeto " & _
                    "From DocVentasDet A " & _
                    "Left Join DocReferencia B " & _
                        "On A.IdEmpresa = B.IdEmpresa And A.IdDocumento = B.TipoDocReferencia And A.IdSerie = B.SerieDocReferencia " & _
                        "And A.IdDocVentas = B.NumDocReferencia " & _
                    "Left Join DocVentas C " & _
                        "On B.IdEmpresa = C.IdEmpresa And B.TipoDocOrigen = C.IdDocumento And B.SerieDocOrigen = C.IdSerie And B.NumDocOrigen = C.IdDocVentas " & _
                        "And 'ANU' <> C.EstDocVentas " & _
                    "Left Join DocVentasDet D " & _
                        "On C.IdEmpresa = D.IdEmpresa And C.IdSucursal = D.IdSucursal And C.IdDocumento = D.IdDocumento And C.IdSerie = D.IdSerie " & _
                        "And C.IdDocVentas = D.IdDocVentas And A.IdProducto = D.IdProducto " & _
                    "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & txtCod_Documento.Text & "' " & _
                    "And A.IdDocVentas = '" & strNumDoc & "' And A.IdSerie = '" & strSerie & "' " & _
                    "Group By A.IdEmpresa,A.IdDocumento,A.IdSerie,A.IdDocVentas,A.IdProducto " & _
                    "Having TotalPVBrutoOri > Sum(IfNull(D.TotalPVBruto,0) * If(IfNull(C.IdDocumento,'') = '07',-1,1))"
        
        Else
            
            If glsValidaStock Then
            
                csql = "Call Spu_ListaDocExportar('" & glsEmpresa & "','','1','" & txtCod_Cliente.Text & "','" & txtCod_Documento.Text & "','" & strSerie & "','" & strNumDoc & "','','" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "','" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "','D')"
            
            Else
            
                csql = "Select Item,IdProducto,IdCodFabricante,GlsProducto,IdMarca,GlsMarca,IdUM,GlsUM,Factor,Afecto," & _
                       "(Cantidad - CantidadImp) As Cantidad,VVUnit,IGVUnit,PVUnit,TotalVVBruto,TotalPVBruto,PorDcto,DctoVV,DctoPV,TotalVVNeto," & _
                       "TotalIGVNeto,TotalPVNeto,IdTipoProducto,IdMoneda,NumLote,FecVencProd,VVUnitLista,PVUnitLista,VVUnitNeto,PVUnitNeto,Cantidad2," & _
                       "CodigoRapido,IdTallaPeso,ItemPro,IvapUnit,TotalIvapNeto " & _
                       "From DocVentasDet " & _
                       "Where IdEmpresa = '" & glsEmpresa & "' And IdSucursal = '" & glsSucursal & "' And IdDocumento = '" & txtCod_Documento.Text & "' " & _
                       "And IdDocVentas = '" & strNumDoc & "' And IdSerie = '" & strSerie & "' And EstDocImportado <> 'S' Order By Item"
            
            End If
            
        End If
                
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
        Do While Not rst.EOF
            RsD.AddNew
            RsD.Fields("Item") = strNumDoc & strSerie & CStr(rst.Fields("Item"))
            RsD.Fields("chkMarca") = 0
            RsD.Fields("idProducto") = rst.Fields("idProducto")
            RsD.Fields("idCodFabricante") = "" & rst.Fields("idCodFabricante")
            RsD.Fields("GlsProducto") = "" & rst.Fields("GlsProducto")
            RsD.Fields("idMarca") = "" & rst.Fields("idMarca")
            RsD.Fields("GlsMarca") = "" & rst.Fields("GlsMarca")
            RsD.Fields("idUM") = "" & rst.Fields("idUM")
            RsD.Fields("GlsUM") = "" & rst.Fields("GlsUM")
            RsD.Fields("Factor") = "" & rst.Fields("Factor")
            RsD.Fields("Afecto") = "" & rst.Fields("Afecto")
            RsD.Fields("Cantidad") = "" & rst.Fields("Cantidad")
            RsD.Fields("Cantidad2") = "" & rst.Fields("Cantidad2")
            RsD.Fields("VVUnit") = "" & rst.Fields("VVUnit")
            RsD.Fields("IGVUnit") = "" & rst.Fields("IGVUnit")
            RsD.Fields("PVUnit") = "" & rst.Fields("PVUnit")
            RsD.Fields("TotalVVBruto") = "" & rst.Fields("TotalVVBruto")
            RsD.Fields("TotalPVBruto") = "" & rst.Fields("TotalPVBruto")
            RsD.Fields("PorDcto") = "" & rst.Fields("PorDcto")
            RsD.Fields("DctoVV") = "" & rst.Fields("DctoVV")
            RsD.Fields("DctoPV") = "" & rst.Fields("DctoPV")
            RsD.Fields("TotalVVNeto") = "" & rst.Fields("TotalVVNeto")
            RsD.Fields("TotalIGVNeto") = "" & rst.Fields("TotalIGVNeto")
            RsD.Fields("TotalPVNeto") = "" & rst.Fields("TotalPVNeto")
            RsD.Fields("idTipoProducto") = "" & rst.Fields("idTipoProducto")
            RsD.Fields("idMoneda") = "" & rst.Fields("idMoneda")
            RsD.Fields("idDocVentas") = strNumDoc
            RsD.Fields("idSerie") = strSerie
            RsD.Fields("NumLote") = "" & rst.Fields("NumLote")
            RsD.Fields("FecVencProd") = "" & rst.Fields("FecVencProd")
            RsD.Fields("VVUnitLista") = "" & rst.Fields("VVUnitLista")
            RsD.Fields("PVUnitLista") = "" & rst.Fields("PVUnitLista")
            RsD.Fields("VVUnitNeto") = "" & rst.Fields("VVUnitNeto")
            RsD.Fields("PVUnitNeto") = "" & rst.Fields("PVUnitNeto")
            RsD.Fields("CodigoRapido") = "" & rst.Fields("CodigoRapido")
            RsD.Fields("idTallaPeso") = Val("" & rst.Fields("idTallaPeso"))
            RsD.Fields("ItemPro") = Val("" & rst.Fields("ItemPro"))
            RsD.Fields("IvapUnit") = Val("" & rst.Fields("IvapUnit"))
            RsD.Fields("TotalIvapNeto") = Val("" & rst.Fields("TotalIvapNeto"))
            rst.MoveNext
        Loop
    End If
    
    If RsD.RecordCount > 0 Then
        mostrarDatosGridSQL gListaDetalle, RsD, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    gListaDetalle.Dataset.Filter = " idDocVentas = '" & strNumDoc & "' AND idSerie = '" & strSerie & "'"
    gListaDetalle.Dataset.Filtered = True
    gListaDetalle.Dataset.EnableControls
    
    Me.Refresh
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub

Err:
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    gListaDetalle.Dataset.EnableControls
    If StrMsgError = "" Then StrMsgError = Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    
    If RsD.State = 1 Then RsD.Close: Set RsD = Nothing

End Sub

Private Sub gLista_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If gLista.Dataset.State = dsEdit Then
        gLista.Dataset.Post
    End If
    If gListaDetalle.Count = 0 Then Exit Sub
    gListaDetalle.Dataset.First
    
    Do While Not gListaDetalle.Dataset.EOF
        gListaDetalle.Dataset.Edit
        gListaDetalle.Columns.ColumnByFieldName("chkMarca").Value = Column.Value
        gListaDetalle.Dataset.Post
        gListaDetalle.Dataset.Next
    Loop

End Sub

Private Sub gListaDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If gListaDetalle.Dataset.State = dsEdit Then
        gListaDetalle.Dataset.Post
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError     As String
Dim strNumDoc       As String
Dim strSerie        As String
Dim PAR             As String

    Select Case Button.Index
        Case 1 'Nuevo
            If gLista.Count > 0 Then
             strNumDoc = gLista.Columns.ColumnByFieldName("idDocVentas").Value
             strSerie = gLista.Columns.ColumnByFieldName("idSerie").Value
             ValCot = traerCampo("docventas", "FecEmision", "iddocventas", strNumDoc, False, " idserie='" & strSerie & "' AND idEmpresa = '" & glsEmpresa & "' AND idSucursal = '" & glsSucursal & "'")
             PAR = traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VALIDEZ_COT_ALSISAC", False, "idEmpresa = '" & glsEmpresa & "'")
             If Trim(txtCod_Documento.Text) = "92" Then
                If Val("" & PAR) > 0 Then
                    If Val(DateDiff("d", CVDate(ValCot), CVDate(getFechaSistema))) >= PAR Then
                           MsgBox ("No se puede importar la Cotización, porque excede su fecha de vigencia.")
                           Exit Sub
                    Else
                        Me.Hide
                    End If
                Else
                    Me.Hide
                End If
             Else
                Me.Hide
             End If
            
            End If
        Case 3
            strRptNum = ""
            strRptSerie = ""
            Me.Hide
          Unload Me
    End Select
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description

MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then gLista.SetFocus

End Sub

Private Sub gLista_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
On Error GoTo Err
Dim StrMsgError As String

    ListaDetalle StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    If gLista.Count > 0 Then
        Me.Hide
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarForm(ByVal strTipoDocQueImporta As String, ByVal strCodCliente As String, ByRef rscd As ADODB.Recordset, ByRef rsdd As ADODB.Recordset, ByRef strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err

    indNuevoDoc = True
    
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    
    strTDExportar = strTipoDocQueImporta
    
    txtCod_Cliente.Text = strCodCliente
    txtCod_Documento.Text = ""
   
    dtp_Desde.Value = getFechaSistema
    dtp_Hasta.Value = dtp_Desde.Value
    
    indNuevoDoc = False
   
    frmListaDocExportar.Show 1
    
    'Quitamos Filtros existentes
    gLista.Dataset.Filter = ""
    gLista.Dataset.Filtered = True
    
    gListaDetalle.Dataset.Filter = ""
    gListaDetalle.Dataset.Filtered = True
    
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    
    If TypeName(rsg) = "Nothing" Then
        Exit Sub
    Else
        If rsg.State = 0 Then
            Exit Sub
        End If
    End If
    
    'Eliminamos los registros q no estan marcados
    If rsg.RecordCount > 0 Then
        rsg.MoveFirst
        Do While Not rsg.EOF
            If rsg.Fields("chkMarca") = "0" Then
                rsg.Delete adAffectCurrent
                rsg.Update
            End If
            rsg.MoveNext
        Loop
    End If
    
    If RsD.RecordCount > 0 Then
        RsD.MoveFirst
        Do While Not RsD.EOF
            If RsD.Fields("chkMarca") = "0" Then
                RsD.Delete adAffectCurrent
                RsD.Update
            End If
            RsD.MoveNext
        Loop
    End If
        
    'Devolvemos valores seleccionados
    strTipoDocImportado = txtCod_Documento.Text
       
    Set rscd = rsg.Clone(adLockReadOnly)
    Set rsdd = RsD.Clone(adLockReadOnly)
    
    
    If rsg.State = 1 Then rsg.Close
    Set rsg = Nothing
    
    If RsD.State = 1 Then RsD.Close
    Set RsD = Nothing

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Cliente_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CLIENTE", txtCod_Cliente, txtGls_Cliente
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Documento_Change()
On Error GoTo Err
Dim StrMsgError As String

    txtGls_Documento.Text = traerCampo("documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)
    
    If indNuevoDoc = False Then
        listaDocVentas StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title

End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DOCUMENOSEXP", txtCod_Documento, txtGls_Documento, " AND c.idDocumento = '" & strTDExportar & "'"
        KeyAscii = 0
    End If

End Sub
