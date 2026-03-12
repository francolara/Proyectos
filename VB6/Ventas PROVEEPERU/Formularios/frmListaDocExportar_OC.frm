VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmListaDocExportar_OC 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Documento"
   ClientHeight    =   8880
   ClientLeft      =   5955
   ClientTop       =   2775
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12540
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
            Picture         =   "frmListaDocExportar_OC.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaDocExportar_OC.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   1164
      ButtonWidth     =   3334
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "             Aceptar             "
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
      ForeColor       =   &H00000000&
      Height          =   8115
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   12510
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   12315
         Begin VB.CommandButton cmbAyudaCliente 
            Height          =   315
            Left            =   6075
            Picture         =   "frmListaDocExportar_OC.frx":3518
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   390
         End
         Begin VB.CommandButton cmbAyudaDocumento 
            Height          =   315
            Left            =   6075
            Picture         =   "frmListaDocExportar_OC.frx":38A2
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   630
            Width           =   390
         End
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   285
            Left            =   600
            TabIndex        =   8
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
            Container       =   "frmListaDocExportar_OC.frx":3C2C
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   990
            TabIndex        =   1
            Top             =   630
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            BackColor       =   16777215
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
            MaxLength       =   8
            Container       =   "frmListaDocExportar_OC.frx":3C48
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   1965
            TabIndex        =   11
            Top             =   630
            Width           =   4080
            _ExtentX        =   7197
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
            Container       =   "frmListaDocExportar_OC.frx":3C64
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txtCod_Cliente 
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            BackColor       =   16777215
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
            MaxLength       =   8
            Container       =   "frmListaDocExportar_OC.frx":3C80
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Cliente 
            Height          =   315
            Left            =   1965
            TabIndex        =   14
            Top             =   240
            Width           =   4080
            _ExtentX        =   7197
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
            Container       =   "frmListaDocExportar_OC.frx":3C9C
            Vacio           =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp_Desde 
            Height          =   315
            Left            =   8085
            TabIndex        =   2
            Top             =   240
            Width           =   1290
            _ExtentX        =   2275
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
            Format          =   132120577
            CurrentDate     =   38955
         End
         Begin MSComCtl2.DTPicker dtp_Hasta 
            Height          =   315
            Left            =   10680
            TabIndex        =   3
            Top             =   240
            Width           =   1290
            _ExtentX        =   2275
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
            Format          =   132120577
            CurrentDate     =   38955
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   7515
            TabIndex        =   17
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   10125
            TabIndex        =   16
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lbl_Cliente 
            Appearance      =   0  'Flat
            Caption         =   "Proveedor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   105
            TabIndex        =   15
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Documento:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   105
            TabIndex        =   12
            Top             =   660
            Width           =   825
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   3825
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportar_OC.frx":3CB8
         TabIndex        =   4
         Top             =   1485
         Width           =   12315
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gListaDetalle 
         Height          =   2535
         Left            =   75
         OleObjectBlob   =   "frmListaDocExportar_OC.frx":8CF2
         TabIndex        =   5
         Top             =   5475
         Width           =   12315
      End
   End
End
Attribute VB_Name = "frmListaDocExportar_OC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsg As New ADODB.Recordset
Private RsD As New ADODB.Recordset
Private strTDExportar As String
Dim indNuevoDoc As Boolean
Dim NIdDPM                              As String

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "PROVEEDOR", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaDocumento_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento, " AND idDocumento in(select idDocumentoExp from documentosexportar where idDocumento = '" & strTDExportar & "')"

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
    
    If strTDExportar = "TE" Then
        txtCod_Documento.Text = "PM"
    Else
        txtCod_Documento.Text = "94"
    End If
    
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

    Set gLista.DataSource = Nothing
    Set rsg = Nothing
    Set RsD = Nothing
    
    rsg.Fields.Append "Item", adChar, 14, adFldRowID
    rsg.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    rsg.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
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
    rsg.Fields.Append "idSucursal", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idAlmacen", adChar, 8, adFldIsNullable
    rsg.Fields.Append "idValescab", adChar, 8, adFldIsNullable
    rsg.Fields.Append "FechaEmisionVale", adVarChar, 10, adFldIsNullable
    rsg.Fields.Append "idPersona", adChar, 8, adFldIsNullable
    rsg.Fields.Append "GlsPersona", adVarChar, 300, adFldIsNullable
    rsg.Fields.Append "ObsDocVentas", adVarChar, 500, adFldIsNullable
    rsg.Fields.Append "IdCentroCosto", adVarChar, 8, adFldIsNullable
    
    If txtCod_Documento.Text = "PM" Then
        
        rsg.Fields.Append "IdUPP", adVarChar, 8, adFldIsNullable
        rsg.Fields.Append "TipoDocReferencia", adVarChar, 2, adFldIsNullable
        rsg.Fields.Append "SerieDocReferencia", adVarChar, 3, adFldIsNullable
        rsg.Fields.Append "NumDocReferencia", adVarChar, 8, adFldIsNullable
        
    End If
    
    rsg.Open , , adOpenKeyset, adLockOptimistic
    
    RsD.Fields.Append "Item", adVarChar, 20, adFldRowID
    RsD.Fields.Append "chkMarca", adChar, 1, adFldIsNullable
    RsD.Fields.Append "idProducto", adChar, 8, adFldIsNullable
    RsD.Fields.Append "CodigoRapido", adVarChar, 20, adFldIsNullable
    RsD.Fields.Append "idCodFabricante", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "GlsProducto", adVarChar, 800, adFldIsNullable
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
    RsD.Fields.Append "idDocumento", adChar, 2, adFldIsNullable
    RsD.Fields.Append "NumLote", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "FecVencProd", adVarChar, 30, adFldIsNullable
    RsD.Fields.Append "VVUnitLista", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "PVUnitLista", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "VVUnitNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "PVUnitNeto", adDouble, 14, adFldIsNullable
    RsD.Fields.Append "idSucursal", adChar, 8, adFldIsNullable
    RsD.Fields.Append "IdDocumentoR", adChar, 2, adFldIsNullable
    RsD.Fields.Append "IdSerieR", adChar, 3, adFldIsNullable
    RsD.Fields.Append "IdDocVentasR", adChar, 8, adFldIsNullable
    
    If txtCod_Documento.Text = "PM" Then
        
        RsD.Fields.Append "StockDisponible", adDouble, 14, adFldIsNullable
        
    End If
    
    RsD.Open , , adOpenKeyset, adLockOptimistic
    
    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " AND GlsVendedor LIKE '%" & strCond & "%'"
    End If
    
    If traerCampo("Documentos", "Aprobacion", "iddocumento", txtCod_Documento.Text, False) = "S" Then
        strFiltroAprob = " AND indAprobado = '1' "
    Else
        strFiltroAprob = ""
    End If
    
    If txtCod_Documento.Text = "40" And traerCampo("Parametros", "ValParametro", "GlsParametro", "APRUEBA_PEDIDO_AUTOMATICO", True) = "1" Then
        csql = "Select ConCat(A.IdDocumento,A.IdDocVentas,A.IdSerie) As Item,A.IdDocumento,A.IdDocVentas,A.IdSerie,A.IdPerVendedor," & _
                "A.GlsVendedor,A.FecEmision,A.EstDocVentas,A.IdMoneda,A.TotalValorVenta,A.TotalIGVVenta,A.TotalPrecioVenta,A.IdSucursal," & _
                "A.IdPerCliente,A.GlsCliente,A.IdAlmacen,A.ObsDocVentas,A.IdCentroCosto " & _
                "From Docventas A " & _
                "Left Join ValesConver B " & _
                    "On A.IdEmpresa = B.IdEmpresa And A.IdCentroCosto = B.IdCentroCosto And (B.EstValeConver <> 'ANU' Or B.EstValeConver Is Null) " & _
                "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' " & _
                "And A.IdDocumento = '" & txtCod_Documento.Text & "' And A.EstDocventas <> 'ANU' " & _
                "And A.IdPerCliente = '" & txtCod_Cliente.Text & "' " & _
                "And A.FecEmision BetWeen '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' And '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' " & _
                "And A.EstDocImportado <> 'S' And (B.IndCierre = '0' Or B.IndCierre Is Null) " & strFiltroAprob
                    
    ElseIf txtCod_Documento.Text = "PM" Then
            
        If leeParametro("IMPORTAR_PM_TRANSFERENCIAS") = "1" Then
        
            csql = "SELECT concat(A.idDocumento,A.idDocVentas,A.idSerie) as Item , A.idDocumento,A.idDocVentas,A.idSerie,A.idPerVendedor,A.GlsVendedor,A.FecEmision," & _
                   "A.estDocVentas,A.idMoneda,A.TotalValorVenta,A.TotalIGVVenta,A.TotalPrecioVenta, A.idSucursal,A.idPerCliente,A.GlsCliente,A.IdAlmacen,A.ObsDocVentas," & _
                   "A.IdCentroCosto,A.IdUPP,B.TipoDocReferencia,B.SerieDocReferencia,B.NumDocReferencia " & _
                   "FROM docventasPedidos A " & _
                   "Inner Join DocReferencia B " & _
                       "On A.IdEmpresa = B.IdEmpresa And A.IdDocumento = B.TipoDocOrigen And A.IdSerie = B.SerieDocOrigen And A.IdDocVentas = B.NumDocOrigen " & _
                   "WHERE A.idEmpresa = '" & glsEmpresa & "' AND A.idSucursal = '" & glsSucursal & "' AND A.idDocumento = '" & txtCod_Documento.Text & "' " & _
                   "AND A.estDocventas <> 'ANU' AND idPerCliente = '" & txtCod_Cliente.Text & "' AND FecEmision between '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' AND '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' AND estDocImportado <> 'S' " & strFiltroAprob
        
        Else
        
            If NIdDPM = "" Then
                NIdDPM = Trim(Format(Now, "ddmmyyy") & Format(Now, "hhMMss") & glsUser & nPC & "_DPM")
            End If
            
            Cn.Execute ("Call Spu_CalculaDespachoPM('" & glsEmpresa & "','" & NIdDPM & "','','0','','','','','','','','','" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "','" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "')")
            
            csql = "Select ConCat(A.IdDocumento,A.IdDocVentas,A.IdSerie) Item,A.IdDocumento,A.IdDocVentas,A.IdSerie,A.IdPerVendedor,A.GlsVendedor,A.FecEmision," & _
                   "A.EstDocVentas,A.IdMoneda,A.TotalValorVenta,A.TotalIGVVenta,A.TotalPrecioVenta,A.IdSucursal,A.IdPerCliente,A.GlsCliente,A.IdAlmacen," & _
                   "A.ObsDocVentas,A.IdCentroCosto,A.IdUPP,B.TipoDocReferencia,B.SerieDocReferencia,B.NumDocReferencia " & _
                   "From DocVentasPedidos A " & _
                   "Inner Join DocReferencia B " & _
                       "On A.IdEmpresa = B.IdEmpresa And A.IdDocumento = B.TipoDocOrigen And A.IdSerie = B.SerieDocOrigen And A.IdDocVentas = B.NumDocOrigen " & _
                   "Inner Join DocVentasPedidosDet C " & _
                       "On A.IdEmpresa = C.IdEmpresa And A.IdSucursal = C.IdSucursal And A.IdDocumento = C.IdDocumento And A.IdSerie = C.IdSerie " & _
                       "And A.IdDocVentas = C.IdDocVentas " & _
                   "Left Join " & NIdDPM & " E " & _
                       "On C.IdEmpresa = E.IdEmpresa And C.IdDocumento = E.IdDocumento And C.IdSerie = E.IdSerie " & _
                       "And C.IdDocVentas = E.IdDocVentas And C.IdProducto = E.IdProducto " & _
                   "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & txtCod_Documento.Text & "' " & _
                   "And A.EstDocVentas <> 'ANU' And A.IdPerCliente = '" & txtCod_Cliente.Text & "' And C.Cantidad > IfNull(E.Despachado,0) " & _
                   "And A.FecEmision BetWeen '" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' AND '" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' " & _
                   "Group By A.IdDocumento,A.IdSerie,A.IdDocVentas"
        
        End If
        
    Else
        'PQS 140616 Ya no filtra Sucursal AND idSucursal = '" & glsSucursal & "'
        
        csql = "SELECT concat(idDocumento,idDocVentas,idSerie) as Item , idDocumento,idDocVentas,idSerie,idPerVendedor,GlsVendedor,FecEmision," & _
               "estDocVentas,idMoneda,TotalValorVenta,TotalIGVVenta,TotalPrecioVenta, idSucursal,idPerCliente,GlsCliente,IdAlmacen,ObsDocVentas," & _
               "IdCentroCosto " & _
               "FROM docventas " & _
               "WHERE idEmpresa = '" & glsEmpresa & "' AND idDocumento = '" & txtCod_Documento.Text & "' " & _
               "AND estDocventas <> 'ANU' AND idPerCliente = '" & txtCod_Cliente.Text & "' " & _
               "AND CAST(FecEmision AS DATE) between CAST('" & Format(dtp_Desde.Value, "yyyy-mm-dd") & "' AS DATE) AND CAST('" & Format(dtp_Hasta.Value, "yyyy-mm-dd") & "' AS DATE) " & _
               "AND estDocImportado <> 'S' And IsNull(IndCerrado,'') <> '1'" & strFiltroAprob
               
    End If
    
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF And Not rst.BOF Then
        Do While Not rst.EOF
            rsg.AddNew
            rsg.Fields("Item") = rst.Fields("Item")
            rsg.Fields("chkMarca") = 0
            rsg.Fields("idDocVentas") = rst.Fields("idDocVentas")
            rsg.Fields("idSerie") = rst.Fields("idSerie")
            rsg.Fields("idDocumento") = rst.Fields("idDocumento")
            rsg.Fields("idPerVendedor") = rst.Fields("idPerVendedor")
            rsg.Fields("GlsVendedor") = rst.Fields("GlsVendedor")
            rsg.Fields("FecEmision") = Format(rst.Fields("FecEmision"), "dd/mm/yyyy")
            rsg.Fields("estDocVentas") = rst.Fields("estDocVentas")
            rsg.Fields("idMoneda") = rst.Fields("idMoneda")
            rsg.Fields("TotalValorVenta") = rst.Fields("TotalValorVenta")
            rsg.Fields("TotalIGVVenta") = rst.Fields("TotalIGVVenta")
            rsg.Fields("TotalPrecioVenta") = rst.Fields("TotalPrecioVenta")
            rsg.Fields("idSucursal") = rst.Fields("idSucursal")
            rsg.Fields("idPersona") = rst.Fields("idPerCliente")
            rsg.Fields("GlsPersona") = rst.Fields("GlsCliente")
            rsg.Fields("IdAlmacen") = rst.Fields("IdAlmacen")
            rsg.Fields("ObsDocVentas") = rst.Fields("ObsDocVentas")
            rsg.Fields("IdCentroCosto") = rst.Fields("IdCentroCosto")
            
            If txtCod_Documento.Text = "PM" Then
            
                rsg.Fields("IdUPP") = rst.Fields("IdUPP")
                rsg.Fields("TipoDocReferencia") = rst.Fields("TipoDocReferencia")
                rsg.Fields("SerieDocReferencia") = rst.Fields("SerieDocReferencia")
                rsg.Fields("NumDocReferencia") = rst.Fields("NumDocReferencia")
                
            End If
            
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
End Sub

Private Sub ListaDetalle(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim strCond As String
Dim indExisteDoc As Boolean
Dim strNumDoc As String
Dim strSerie  As String
Dim strTD     As String

    strNumDoc = gLista.Columns.ColumnByFieldName("idDocVentas").Value
    strSerie = gLista.Columns.ColumnByFieldName("idSerie").Value
    strTD = gLista.Columns.ColumnByFieldName("idDocumento").Value
    
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
        strCond = ""
        If Trim(txt_TextoBuscar.Text) <> "" Then
            strCond = Trim(txt_TextoBuscar.Text)
            strCond = " AND GlsCliente LIKE '%" & strCond & "%'"
        End If
        
        If txtCod_Documento.Text = "PM" Then
            
            If leeParametro("IMPORTAR_PM_TRANSFERENCIAS") = "1" Then
                
                csql = "SELECT A.item, A.idProducto,B.CodigoRapido,A.idCodFabricante, A.GlsProducto,A.idMarca, A.GlsMarca,A.idUM, A.GlsUM,A.Factor,A.Afecto," & _
                       "A.Cantidad,A.VVUnit,A.IGVUnit, A.PVUnit,A.TotalVVBruto,A.TotalPVBruto, A.PorDcto,A.DctoVV,A.DctoPV," & _
                       "A.TotalVVNeto,A.TotalIGVNeto, A.TotalPVNeto,A.idTipoProducto,A.idMoneda,A.NumLote,A.FecVencProd,A.VVUnitLista,A.PVUnitLista," & _
                       "A.VVUnitNeto, A.PVUnitNeto,A.Cantidad2, A.idSucursal,A.IdDocumentoImp IdDocumentoR,A.IdSerieImp IdSerieR," & _
                       "A.IdDocVentasImp IdDocVentasR " & _
                       "From DocVentasPedidosDet A " & _
                       "Inner Join Productos B " & _
                           "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                       "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & txtCod_Documento.Text & "' " & _
                       "AND A.idDocVentas = '" & strNumDoc & "' AND A.idSerie = '" & strSerie & "'"
                       
            Else
            
                csql = "SELECT A.item, A.idProducto,B.CodigoRapido,A.idCodFabricante, A.GlsProducto,A.idMarca, A.GlsMarca,A.idUM, A.GlsUM,A.Factor,A.Afecto," & _
                       "(A.Cantidad - IfNull(E.Despachado,0)) as Cantidad,A.VVUnit,A.IGVUnit, A.PVUnit,A.TotalVVBruto,A.TotalPVBruto, A.PorDcto,A.DctoVV,A.DctoPV," & _
                       "A.TotalVVNeto,A.TotalIGVNeto, A.TotalPVNeto,A.idTipoProducto,A.idMoneda,A.NumLote,A.FecVencProd,A.VVUnitLista,A.PVUnitLista," & _
                       "A.VVUnitNeto, A.PVUnitNeto,A.Cantidad2, A.idSucursal,A.IdDocumentoImp IdDocumentoR,A.IdSerieImp IdSerieR," & _
                       "A.IdDocVentasImp IdDocVentasR " & _
                       "From DocVentasPedidosDet A " & _
                       "Inner Join Productos B " & _
                           "On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                       "Left Join " & NIdDPM & " E " & _
                           "On A.IdEmpresa = E.IdEmpresa And A.IdDocumento = E.IdDocumento And A.IdSerie = E.IdSerie And A.IdDocVentas = E.IdDocVentas " & _
                           "And A.IdProducto = E.IdProducto " & _
                       "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdSucursal = '" & glsSucursal & "' And A.IdDocumento = '" & txtCod_Documento.Text & "' " & _
                       "AND A.idDocVentas = '" & strNumDoc & "' AND A.idSerie = '" & strSerie & "' And (A.Cantidad - IfNull(E.Despachado,0)) > 0"
            
            End If
            
        Else
        
            csql = "SELECT A.item, A.idProducto,B.CodigoRapido,A.idCodFabricante, A.GlsProducto,A.idMarca, A.GlsMarca,A.idUM, A.GlsUM,A.Factor,A.Afecto," & _
                   "(A.Cantidad - A.CantidadImp) as Cantidad,A.VVUnit,A.IGVUnit, A.PVUnit,A.TotalVVBruto,A.TotalPVBruto, A.PorDcto,A.DctoVV,A.DctoPV," & _
                   "A.TotalVVNeto,A.TotalIGVNeto, A.TotalPVNeto,A.idTipoProducto,A.idMoneda,A.NumLote,A.FecVencProd,A.VVUnitLista,A.PVUnitLista," & _
                   "A.VVUnitNeto, A.PVUnitNeto,A.Cantidad2, A.idSucursal,A.IdDocumentoImp IdDocumentoR,A.IdSerieImp IdSerieR," & _
                   "A.IdDocVentasImp IdDocVentasR " & _
                   "FROM docventasdet A Inner Join Productos B On A.IdEmpresa = B.IdEmpresa And A.IdProducto = B.IdProducto " & _
                   "WHERE A.idEmpresa = '" & glsEmpresa & "' AND A.idDocumento = '" & txtCod_Documento.Text & "' AND A.idDocVentas = '" & strNumDoc & "' AND A.idSerie = '" & strSerie & "' AND A.estDocImportado <> 'S'"
        
        End If
        
        rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            RsD.AddNew
            RsD.Fields("Item") = strNumDoc & strSerie & CStr(rst.Fields("Item"))
            RsD.Fields("chkMarca") = 0
            RsD.Fields("idProducto") = rst.Fields("idProducto")
            RsD.Fields("CodigoRapido") = "" & rst.Fields("CodigoRapido")
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
            RsD.Fields("VVUnit") = "" & rst.Fields("VVUnitNeto")
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
            RsD.Fields("idDocumento") = strTD
            RsD.Fields("NumLote") = "" & rst.Fields("NumLote")
            RsD.Fields("FecVencProd") = "" & rst.Fields("FecVencProd")
            RsD.Fields("VVUnitLista") = "" & rst.Fields("VVUnitLista")
            RsD.Fields("PVUnitLista") = "" & rst.Fields("PVUnitLista")
            RsD.Fields("VVUnitNeto") = "" & rst.Fields("VVUnitNeto")
            RsD.Fields("PVUnitNeto") = "" & rst.Fields("PVUnitNeto")
            RsD.Fields("idSucursal") = "" & rst.Fields("idSucursal")
            RsD.Fields("IdDocumentoR") = "" & rst.Fields("IdDocumentoR")
            RsD.Fields("IdSerieR") = "" & rst.Fields("IdSerieR")
            RsD.Fields("IdDocVentasR") = "" & rst.Fields("IdDocVentasR")
            
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
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    gListaDetalle.Dataset.EnableControls
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    If RsD.State = 1 Then RsD.Close: Set RsD = Nothing

End Sub

Private Sub gLista_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
On Error GoTo Err
Dim rsLista             As New ADODB.Recordset
Dim RsListaClone        As New ADODB.Recordset
Dim StrMsgError         As String
    
    If gLista.Dataset.State = dsEdit Then
        gLista.Dataset.Post
    End If
    
    If strTDExportar = "99" And traerCampo("Parametros", "ValParametro", "GlsParametro", "APRUEBA_PEDIDO_AUTOMATICO", True) = "1" Then
        Set rsLista = gLista.DataSource
        Set RsListaClone = rsLista.Clone(adLockOptimistic)
        
        RsListaClone.Filter = "chkMarca = '1'"
        If RsListaClone.RecordCount = 2 Then
            gLista.Dataset.Edit
            gLista.Columns.ColumnByFieldName("ChkMarca").Value = "0"
            gLista.Dataset.Post
            RsListaClone.Filter = ""
            RsListaClone.Close
            StrMsgError = "Sólo puede importar un Pedido por Conversión": GoTo Err
        End If
        RsListaClone.Filter = ""
        RsListaClone.Close
    End If
    
    If gListaDetalle.Count = 0 Then Exit Sub
    
    gListaDetalle.Dataset.First
    Do While Not gListaDetalle.Dataset.EOF
        gListaDetalle.Dataset.Edit
        gListaDetalle.Columns.ColumnByFieldName("chkMarca").Value = Column.Value
        gListaDetalle.Dataset.Post
        gListaDetalle.Dataset.Next
    Loop
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub gListaDetalle_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    
    If gListaDetalle.Dataset.State = dsEdit Then
        gListaDetalle.Dataset.Post
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String
    
    Select Case Button.Index
        Case 1 'Nuevo
            If gLista.Count > 0 Then
                Me.Hide
            End If
        Case 3
            strRptNum = ""
            strRptSerie = ""
            If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
            Me.Hide
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
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub MostrarForm(ByVal strTipoDocQueImporta As String, ByVal strCodCliente As String, ByRef rscd As ADODB.Recordset, ByRef rsdd As ADODB.Recordset, ByRef strTipoDocImportado As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsa As New ADODB.Recordset

    indNuevoDoc = True
    
    strTDExportar = strTipoDocQueImporta
    
    Set gLista.DataSource = Nothing
    Set gListaDetalle.DataSource = Nothing
    
    txtCod_Cliente.Text = strCodCliente
    
    csql = "select idDocumento, idDocumentoExp, item from documentosexportar where idDocumento = '" & strTDExportar & "' "
    If rsa.State = 1 Then rsa.Close
    rsa.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    
    If rsa.RecordCount = 1 Then
        txtCod_Documento.Text = "" & rsa.Fields("idDocumentoExp")
    Else
        txtCod_Documento.Text = ""
    End If
   
    dtp_Hasta.Value = getFechaSistema
    dtp_Desde.Value = "01/" & Format(Month(dtp_Hasta.Value), "00") & "/" & Format(Year(dtp_Hasta.Value), "0000")
    indNuevoDoc = False
    
    listaDocVentas StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    frmListaDocExportar_OC.Show 1
    
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
    
    If rsg.State = 1 Then rsg.Close: Set rsg = Nothing
    If RsD.State = 1 Then RsD.Close: Set RsD = Nothing

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
