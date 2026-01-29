VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMantTiposCambio 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tipos de Cambio"
   ClientHeight    =   8820
   ClientLeft      =   2925
   ClientTop       =   1350
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   7470
      Top             =   540
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
            Picture         =   "frmMantTiposCambio.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":07EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":0B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":12BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":2122
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":24BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantTiposCambio.frx":317E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   8115
      Left            =   30
      TabIndex        =   9
      Top             =   645
      Width           =   9330
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   9075
         Begin CATControls.CATTextBox txt_TextoBuscar 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   255
            Width           =   7950
            _ExtentX        =   14023
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
            MaxLength       =   255
            Container       =   "frmMantTiposCambio.frx":3518
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Búsqueda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   11
            Top             =   315
            Width           =   735
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   6915
         Left            =   120
         OleObjectBlob   =   "frmMantTiposCambio.frx":3534
         TabIndex        =   7
         Top             =   1050
         Width           =   9075
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8130
      Left            =   45
      TabIndex        =   8
      Top             =   645
      Width           =   9315
      Begin VB.Frame FraTCSBS 
         Caption         =   " Tipo de Cambio SBS "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   1035
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   6500
         Begin CATControls.CATTextBox txtTc_CompraSBS 
            Height          =   315
            Left            =   1350
            TabIndex        =   5
            Tag             =   "NtcCompraSBS"
            Top             =   480
            Width           =   840
            _ExtentX        =   1482
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantTiposCambio.frx":558C
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtTc_VentaSBS 
            Height          =   315
            Left            =   4815
            TabIndex        =   6
            Tag             =   "NtcVentaSBS"
            Top             =   480
            Width           =   840
            _ExtentX        =   1482
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantTiposCambio.frx":55A8
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4230
            TabIndex        =   23
            Top             =   555
            Width           =   435
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   675
            TabIndex        =   22
            Top             =   550
            Width           =   555
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Tipo de Cambio Comercial "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   1035
         TabIndex        =   16
         Top             =   3060
         Width           =   6500
         Begin CATControls.CATTextBox CATTextBox2 
            Height          =   315
            Left            =   4815
            TabIndex        =   4
            Tag             =   "NTcFacturacion"
            Top             =   495
            Width           =   840
            _ExtentX        =   1482
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantTiposCambio.frx":55C4
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txt_tcCompraComercial 
            Height          =   315
            Left            =   1350
            TabIndex        =   3
            Tag             =   "NtcCompraC"
            Top             =   495
            Width           =   840
            _ExtentX        =   1482
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantTiposCambio.frx":55E0
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   675
            TabIndex        =   20
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4230
            TabIndex        =   19
            Top             =   540
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Cambio SUNAT "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   1035
         TabIndex        =   15
         Top             =   1800
         Width           =   6500
         Begin CATControls.CATTextBox txtVal_VV 
            Height          =   315
            Left            =   1350
            TabIndex        =   1
            Tag             =   "NTcCompra"
            Top             =   480
            Width           =   840
            _ExtentX        =   1482
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantTiposCambio.frx":55FC
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox CATTextBox1 
            Height          =   315
            Left            =   4815
            TabIndex        =   2
            Tag             =   "NTcVenta"
            Top             =   480
            Width           =   840
            _ExtentX        =   1482
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
            Alignment       =   1
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmMantTiposCambio.frx":5618
            Text            =   "0.000"
            Decimales       =   3
            Estilo          =   4
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Compra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   675
            TabIndex        =   18
            Top             =   550
            Width           =   555
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4230
            TabIndex        =   17
            Top             =   555
            Width           =   435
         End
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1845
         TabIndex        =   13
         Tag             =   "FFecha"
         Top             =   1125
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
         Format          =   62062593
         CurrentDate     =   38638
      End
      Begin VB.Label lblFecNacimiento 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1080
         TabIndex        =   14
         Top             =   1215
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   2170
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgDocVentas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "         Nuevo         "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMantTiposCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String

    ConfGrid gLista, False, False, False, False
    listaTC StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    fraListado.Visible = True
    fraGeneral.Visible = False
    habilitaBotones 7
    nuevo
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "TIPO_CAMBIO_BANCOS", True) = "S" Then
        FraTCSBS.Visible = True
    Else
        FraTCSBS.Visible = False
    End If
    
    
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Grabar(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCodigo   As String
Dim strMsg      As String, cselect As String
Dim rst         As New ADODB.Recordset
Dim sw          As Boolean
Dim NMontoTC    As Integer
    
    NMontoTC = Val("" & traerCampo("Parametros", "ValParametro", "GlsParametro", "MONTO_MAXIMO_TIPO_DE_CAMBIO", True))
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    If Val("" & txtVal_VV.Text) > NMontoTC Then
        StrMsgError = "El Monto ingresado para el T/C '" & Label2.Caption & "' es mayor al Monto Máximo  de Tipo de Cambio"
        txtVal_VV.SetFocus
        GoTo Err
        
    ElseIf Val("" & CATTextBox1.Text) > NMontoTC Then
        StrMsgError = "El Monto ingresado para el T/C  '" & Label1.Caption & "' es mayor al Monto Máximo  de Tipo de Cambio"
        CATTextBox1.SetFocus
        GoTo Err
    
    ElseIf Val("" & CATTextBox2.Text) > FormatNumber(NMontoTC, glsDecimalesTC) Then
        StrMsgError = "El Monto ingresado para el T/C   '" & Label4.Caption & "' es mayor al Monto Máximo  de Tipo de Cambio"
        CATTextBox2.SetFocus
        GoTo Err
        
    ElseIf Val("" & txt_tcCompraComercial.Text) > FormatNumber(NMontoTC, glsDecimalesTC) Then
        StrMsgError = "El Monto ingresado para el T/C   '" & Label5.Caption & "' es mayor al Monto Máximo  de Tipo de Cambio"
         txt_tcCompraComercial.SetFocus
        GoTo Err
    End If
    
    cselect = "SELECT FECHA FROM tiposdecambio WHERE YEAR(FECHA) = '" & Year(DtpFecha.Value) & "' and month(fecha) = '" & Month(DtpFecha.Value) & "' and day(fecha) = '" & Day(DtpFecha.Value) & "'"
    If rst.State = 1 Then rst.Close
    rst.Open cselect, Cn, adOpenStatic, adLockReadOnly
    If rst.RecordCount = 0 Then
        sw = True
    Else
        sw = False
    End If
    
    If sw = True Then '--- Graba
        EjecutaSQLForm Me, 0, False, "tiposdecambio", StrMsgError
        If StrMsgError <> "" Then GoTo Err
        strMsg = "Grabo"
    
    Else '--- Modifica
        csql = "update tiposdecambio set tcFacturacion = " & Val(CATTextBox2.Text) & ", " & _
               "tcVenta = " & Val(CATTextBox1.Text) & ", " & _
               "tcCompra = " & Val(txtVal_VV.Text) & ", " & _
               "tcCompraC = " & Val(txt_tcCompraComercial.Text) & ", " & _
               "tcCompraSBS = " & Val(txtTc_CompraSBS.Text) & ", " & _
               "tcVentaSBS = " & Val(txtTc_VentaSBS.Text) & " " & _
               "where year(FECHA) = '" & Year(DtpFecha.Value) & "' " & _
               "and month(fecha) = '" & Month(DtpFecha.Value) & "' " & _
               "and day(fecha) = '" & Day(DtpFecha.Value) & "' "
        Cn.Execute csql
        
        strMsg = "Modifico"
    End If
    MsgBox "Se " & strMsg & " Satisfactoriamente", vbInformation, App.Title
    
    fraGeneral.Enabled = False
    listaTC StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub nuevo()

    limpiaForm Me
    
End Sub

Private Sub gLista_OnDblClick()
On Error GoTo Err
Dim StrMsgError As String

    mostrarTC Format(gLista.Columns.ColumnByName("Fecha").Value, "dd/mm/yyyy"), StrMsgError
    If StrMsgError <> "" Then GoTo Err
    fraListado.Visible = False
    fraGeneral.Visible = True
    fraGeneral.Enabled = False
    habilitaBotones 2
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
Dim StrMsgError As String

    Select Case Button.Index
        Case 1 'Nuevo
            nuevo
            fraListado.Visible = False
            fraGeneral.Visible = True
            fraGeneral.Enabled = True
        Case 2 'Grabar
            Grabar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 3 'Modificar
            fraGeneral.Enabled = True
        Case 4, 7  'Cancelar
            fraListado.Visible = True
            fraGeneral.Visible = False
            fraGeneral.Enabled = False
        Case 5 'Eliminar
            eliminar StrMsgError
            If StrMsgError <> "" Then GoTo Err
        Case 6 'Imprimir
        
            gLista.m.ExportToXLS App.Path & "\Temporales\TiposdeCambio.xls"
            ShellEx App.Path & "\Temporales\TiposdeCambio.XLS", essSW_MAXIMIZE, , , "open", Me.hwnd
            If Len(Trim(StrMsgError)) = 0 Then
                MsgBox "Fin de Proceso", vbInformation, "Ventas"
            End If
        Case 8 'Salir
            Unload Me
    End Select
    habilitaBotones Button.Index
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub habilitaBotones(indexBoton As Integer)
Dim indHabilitar As Boolean

    Select Case indexBoton
        Case 1, 2, 3 'Nuevo, Grabar, Modificar
            If indexBoton = 2 Then indHabilitar = True
            Toolbar1.Buttons(1).Visible = indHabilitar 'Nuevo
            Toolbar1.Buttons(2).Visible = Not indHabilitar 'Grabar
            Toolbar1.Buttons(3).Visible = indHabilitar 'Modificar
            Toolbar1.Buttons(4).Visible = Not indHabilitar 'Cancelar
            Toolbar1.Buttons(5).Visible = indHabilitar 'Eliminar
            Toolbar1.Buttons(6).Visible = False 'Imprimir
            Toolbar1.Buttons(7).Visible = indHabilitar 'Lista
        Case 4, 7 'Cancelar, Lista
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
    End Select

End Sub

Private Sub txt_TextoBuscar_Change()
On Error GoTo Err
Dim StrMsgError As String

    listaTC StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txt_TextoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then gLista.SetFocus
    
End Sub

Private Sub listaTC(ByRef StrMsgError As String)
On Error GoTo Err
Dim strCond As String
Dim rsdatos                     As New ADODB.Recordset

    strCond = ""
    If Trim(txt_TextoBuscar.Text) <> "" Then
        strCond = Trim(txt_TextoBuscar.Text)
        strCond = " WHERE a.Fecha LIKE '%" & strCond & "%'"
    End If
    
    csql = "SELECT CAST(a.Fecha AS DATE) as Fecha, a.TCCompra, a.TcVenta, a.TCFacturacion " & _
           "FROM tiposdecambio a "
    If strCond <> "" Then csql = csql & strCond

    csql = csql & " ORDER BY A.FECHA DESC"
    
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set gLista.DataSource = rsdatos

'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = csql
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "Fecha"
'    End With
    
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub mostrarTC(strCod As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset

    csql = "SELECT a.* " & _
           "FROM tiposdecambio a " & _
           "WHERE YEAR(A.FECHA) = '" & right(strCod, 4) & "' AND MONTH(A.FECHA) = '" & Mid(strCod, 4, 2) & "' AND DAY(A.FECHA) = '" & left(strCod, 2) & "'"
    If rst.State = adStateOpen Then rst.Close
    rst.Open csql, Cn, adOpenStatic, adLockReadOnly
    mostrarDatosFormSQL Me, rst, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Refresh
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub eliminar(ByRef StrMsgError As String)
On Error GoTo Err
Dim indTrans As Boolean
Dim strCodigo As String
Dim rsValida As New ADODB.Recordset

    If MsgBox("¿Seguro de eliminar el registro?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    strCodigo = Trim(DtpFecha.Value)
    
    Cn.BeginTrans
    indTrans = True
    
    '--- Eliminando el registro
    csql = "DELETE FROM tiposdecambio WHERE DATE_FORMAT(FECHA,'%d/%m/%Y') = DATE_FORMAT('" & Format(strCodigo, "yyyy-mm-dd") & "','%d/%m/%Y')"
    
    Cn.Execute csql
    
    Cn.CommitTrans
    
    '--- Nuevo
    MsgBox "Registro eliminado satisfactoriamente", vbInformation, App.Title
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    
    Exit Sub
    
Err:
    If rsValida.State = 1 Then rsValida.Close: Set rsValida = Nothing
    If indTrans Then Cn.RollbackTrans
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtVal_VV_LostFocus()
    If Val(txt_tcCompraComercial.Text) = 0 Then
        txt_tcCompraComercial.Text = txtVal_VV.Text
    End If
End Sub

Private Sub CATTextBox1_LostFocus()
    If Val(CATTextBox2.Text) = 0 Then
        CATTextBox2.Text = CATTextBox1.Text
    End If
End Sub




























