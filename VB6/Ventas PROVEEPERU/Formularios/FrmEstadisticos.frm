VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmEstadisticos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadístico de Ventas"
   ClientHeight    =   5115
   ClientLeft      =   3975
   ClientTop       =   1590
   ClientWidth     =   7425
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4545
      Width           =   1185
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4545
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   7260
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
         Left            =   6650
         Picture         =   "FrmEstadisticos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3780
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
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
         Left            =   6650
         Picture         =   "FrmEstadisticos.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3375
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaVendedor 
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
         Left            =   6650
         Picture         =   "FrmEstadisticos.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1530
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaSucursal 
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
         Left            =   6650
         Picture         =   "FrmEstadisticos.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1170
         Width           =   390
      End
      Begin VB.Frame Frame2 
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
         Height          =   810
         Left            =   180
         TabIndex        =   14
         Top             =   225
         Width           =   6870
         Begin VB.ComboBox CboTipoReporte 
            Height          =   330
            ItemData        =   "FrmEstadisticos.frx":0E28
            Left            =   2295
            List            =   "FrmEstadisticos.frx":0E38
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   270
            Width           =   3705
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Reporte"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   225
            TabIndex        =   15
            Top             =   315
            Width           =   1140
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   2025
         Width           =   6870
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1575
            TabIndex        =   3
            Top             =   345
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   104792065
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4875
            TabIndex        =   4
            Top             =   360
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   104792065
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   13
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4365
            TabIndex        =   12
            Top             =   405
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1100
         TabIndex        =   1
         Tag             =   "TidMoneda"
         Top             =   1185
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "FrmEstadisticos.frx":0F87
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   1185
         Width           =   4580
         _ExtentX        =   8070
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
         Container       =   "FrmEstadisticos.frx":0FA3
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   1100
         TabIndex        =   2
         Tag             =   "TidPerCliente"
         Top             =   1560
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "FrmEstadisticos.frx":0FBF
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Tag             =   "TGlsCliente"
         Top             =   1560
         Width           =   4580
         _ExtentX        =   8070
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
         Locked          =   -1  'True
         Container       =   "FrmEstadisticos.frx":0FDB
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   1095
         TabIndex        =   5
         Tag             =   "TidSerie"
         Top             =   3060
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Arial"
         FontSize        =   9.75
         ForeColor       =   -2147483640
         MaxLength       =   3
         Container       =   "FrmEstadisticos.frx":0FF7
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1095
         TabIndex        =   6
         Tag             =   "TidMoneda"
         Top             =   3405
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "FrmEstadisticos.frx":1013
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Top             =   3405
         Width           =   4575
         _ExtentX        =   8070
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
         Container       =   "FrmEstadisticos.frx":102F
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1095
         TabIndex        =   7
         Tag             =   "TidPerCliente"
         Top             =   3780
         Width           =   915
         _ExtentX        =   1614
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
         Container       =   "FrmEstadisticos.frx":104B
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Tag             =   "TGlsCliente"
         Top             =   3780
         Width           =   4575
         _ExtentX        =   8070
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
         Locked          =   -1  'True
         Container       =   "FrmEstadisticos.frx":1067
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label lblcliente 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   28
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   25
         Top             =   3480
         Width           =   570
      End
      Begin VB.Label lbl_Serie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   22
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   21
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   18
         Top             =   1230
         Width           =   645
      End
   End
   Begin MSComctlLib.ImageList imgDocVentas 
      Left            =   7695
      Top             =   450
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
            Picture         =   "FrmEstadisticos.frx":1083
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":141D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":186F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":1C09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":1FA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":233D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":26D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":2A71
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":2E0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":31A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":353F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadisticos.frx":4201
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmEstadisticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboTipoReporte_Click()

    If right(CboTipoReporte.Text, 1) = "3" Then
        lblcliente.Enabled = True
        txtCod_Cliente.Enabled = True
        txtGls_Cliente.Enabled = True
        cmbAyudaCliente.Enabled = True
    Else
        lblcliente.Enabled = False
        txtCod_Cliente.Enabled = False
        txtGls_Cliente.Enabled = False
        cmbAyudaCliente.Enabled = False
    End If

End Sub

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaVendedor_Click()
   
   mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim GlsReporte As String
Dim GlsForm As String
    
    If right(CboTipoReporte.Text, 1) = "1" Then
        If sw_estadistico = True Then
            GlsReporte = "rptEstadisticoporProductoEmpresasRelacionadas.rpt"
            GlsForm = "Estadistico por Producto - Empresas Relacionadas"
        Else
            GlsReporte = "rptEstadisticoporProducto.rpt"
            GlsForm = "Estadistico por Producto"
        End If
    ElseIf right(CboTipoReporte.Text, 1) = "2" Then
        If sw_estadistico = True Then
            GlsReporte = "rptEstadisticoporCanalEmpresasRelacionadas.rpt"
            GlsForm = "Estadistico por Canal - Empresas Relacionadas"
        Else
            GlsReporte = "rptEstadisticoporCanal.rpt"
            GlsForm = "Estadistico por Canal"
        End If
    ElseIf right(CboTipoReporte.Text, 1) = "3" Then
        If sw_estadistico = True Then
            GlsReporte = "rptEstadisticoporClienteEmpresasRelacionadas.rpt"
            GlsForm = "Estadistico por Cliente - Empresas Relacionadas"
        Else
            GlsReporte = "rptEstadisticoporCliente.rpt"
            GlsForm = "Estadistico por Cliente"
        End If
    ElseIf right(CboTipoReporte.Text, 1) = "4" Then
        GlsReporte = "rptEstadisticoporCreditoPromedio.rpt"
        GlsForm = "Estadistico por Credito Promedio"
    End If
    
    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    imprimir right(CboTipoReporte.Text, 1), GlsReporte, GlsForm, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0
    If sw_estadistico = 0 Then
        Me.Caption = "Estadístico de Ventas"
    Else
        Me.Caption = "Estadístico de Ventas - Empresas Relacionadas"
    End If
        
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    CboTipoReporte.ListIndex = 0
    
    txtCod_Moneda.Text = "PEN"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"

End Sub

Private Sub txtCod_Cliente_Change()
 
    If txtCod_Cliente.Text <> "" Then
        If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
            If Trim(traerCampo("Vendedores", "indVisualizaclientes", "idVendedor", glsUser, False)) = 0 Then
                txtGls_Cliente.Text = traerCampo("personas p Inner Join  clientes c On  p.idPersona = c.idCliente Inner Join personas v On c.idVendedorCampo = v.idPersona", "p.GlsPersona", "p.idPersona", txtCod_Cliente.Text, False, "c.idVendedorCampo ='" & glsUser & "' ")
            Else
                txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
            End If
        Else
            txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
        End If
    Else
        txtGls_Cliente.Text = ""
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text = "" Then
        txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    Else
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    End If

End Sub

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text = "" Then
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    Else
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    End If

End Sub

Private Sub imprimir(ByRef Report As String, ByRef GlsReporte As String, ByRef GlsForm As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim rsReporte       As New ADODB.Recordset
Dim rsEstadistico   As New ADODB.Recordset
Dim rsFamilia       As New ADODB.Recordset
Dim codFamilias     As String
Dim sucursal        As String
Dim Vendedor        As String
Dim serie           As String
Dim empresa         As String
Dim cempresas       As String
    
    empresa = traerCampo("Empresas", "glsEmpresa", "idEmpresa", glsEmpresa, False)
    
    If Trim(txtCod_Sucursal.Text) = "" Then
        sucursal = "TODAS LAS SUCURSALES"
    Else
        sucursal = traerCampo("Personas", "glsPersona", "idPersona", Trim(txtCod_Sucursal.Text), False)
    End If
    
    If Trim(txt_serie.Text) = "" Then
        serie = "TODAS LAS SERIES"
    Else
        serie = Format(Trim(txt_serie.Text), "000")
    End If
    
    If Trim(txtCod_Vendedor.Text) = "" Then
        Vendedor = "TODOS LOS VENDEDORES"
    Else
        Vendedor = traerCampo("Personas", "glsPersona", "idPersona", Trim(txtCod_Vendedor.Text), False)
    End If
    
    csql = "delete from tempEstadistico"
    Cn.Execute csql
    
    '-----------------------------------------------------------------------------------
    '---- CONDICION PARA SELECCIONAR DE EMPRESAS RELACIONADAS
    cempresas = ""
    If sw_estadistico = True Then
        cempresas = "inner join empresasrelacionadas er on a.idempresa = er.idempresa and a.idpercliente = er.idpersona "
    End If
    '-----------------------------------------------------------------------------------
    
    If Report = "1" Then
        csql = "insert into tempEstadistico(IDPRODUCTO, GLSPRODUCTO, CANTIDAD, CANTIDAD2, VVUNIT, PVUNIT, IDMONEDA, TIPOCAMBIO, IDPERCLIENTE, GLSCLIENTE, IDNIVEL, GLSNIVEL, IDFAMILIA, GLSFAMILIA, IDDOCUMENTO, EMPRESA, SUCURSAL, VENDEDOR, SERIE) " & _
                "Select b.idProducto, c.glsProducto, b.Cantidad, b.Cantidad2, " & _
                "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  IF(a.idMoneda = 'PEN', b.TotalVVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalVVNeto * tc.tcVenta)) " & _
                                                        "WHEN 'USD' THEN  IF(a.idMoneda = 'USD', b.TotalVVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalVVNeto / tc.tcVenta)) end," & _
                "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  IF(a.idMoneda = 'PEN', b.TotalPVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalPVNeto * tc.tcVenta)) " & _
                                                        "WHEN 'USD' THEN  IF(a.idMoneda = 'USD', b.TotalPVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalPVNeto / tc.tcVenta)) end," & _
                "'" & Trim(txtCod_Moneda.Text) & "', isnull(tc.tcVenta) as TipoCambio, a.idPerCliente, a.glsCliente, c.idNivel, d.glsNivel, e.idNivel as idFamilia, e.glsNivel as glsFamilia, a.idDocumento, '" & empresa & "', '" & sucursal & "', '" & Vendedor & "', '" & serie & "' " & _
                "from docventas a inner join docventasdet b   " & _
                "on a.idDocVentas = b.idDocVentas and a.idSerie = b.idSerie and a.idDocumento = b.idDocumento and a.idEmpresa = b.idEmpresa and a.idSucursal = b.idSucursal " & _
                cempresas & _
                "inner join Productos c on b.idProducto = c.idProducto and b.idEmpresa = c.idEmpresa " & _
                "left join Niveles d on d.idNivel = c.idNivel and d.idEmpresa = a.idEmpresa left join Niveles e on d.idNivelPred = e.idNivel and d.idEmpresa = e.idEmpresa " & _
                "left join TiposdeCambio tc on year(tc.fecha) = year(a.FecEmision) and month(tc.fecha) = month(a.FecEmision) and day(tc.fecha) = day(a.FecEmision) " & _
                "where a.estDocventas in ('IMP','CAN') and a.idDocumento in('01','03','25') " & _
                "and a.FecEmision between cast('" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "' as date) and cast('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' as date) " & _
                "and a.idSucursal like '%" & Trim(txtCod_Sucursal.Text) & "%' and a.idPerVendedorCampo like '%" & Trim(txtCod_Vendedor.Text) & "%' " & _
                "and a.idSerie like '%" & Format(Trim(txt_serie.Text), "000") & "%' and a.idEmpresa = '" & glsEmpresa & "' "
        Cn.Execute csql
    
        mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parReporte|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Report & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    
    ElseIf Report = "2" Then
        
        mostrarReporte GlsReporte, "parEmpresa|parSucursal|parSerie|parVendedor|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Format(Trim(txt_serie.Text), "000") & "|" & Trim(txtCod_Vendedor.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
    ElseIf Report = "3" Then
        
        csql = "insert into tempEstadistico(IDPRODUCTO, GLSPRODUCTO, CANTIDAD, CANTIDAD2, VVUNIT, PVUNIT, IDMONEDA, TIPOCAMBIO, IDPERCLIENTE, GLSCLIENTE, IDNIVEL, GLSNIVEL, IDFAMILIA, GLSFAMILIA, IDDOCUMENTO, EMPRESA, SUCURSAL, VENDEDOR, SERIE)" & _
                "Select b.idProducto, c.glsProducto, b.Cantidad, b.Cantidad2, " & _
                "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  IF(a.idMoneda = 'PEN', b.TotalVVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalVVNeto * tc.tcVenta)) " & _
                                                        "WHEN 'USD' THEN  IF(a.idMoneda = 'USD', b.TotalVVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalVVNeto / tc.tcVenta)) end," & _
                "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  IF(a.idMoneda = 'PEN', b.TotalPVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalPVNeto * tc.tcVenta)) " & _
                                                        "WHEN 'USD' THEN  IF(a.idMoneda = 'USD', b.TotalPVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalPVNeto / tc.tcVenta)) end," & _
                "'" & Trim(txtCod_Moneda.Text) & "', isnull(tc.tcVenta) as TipoCambio, a.idPerCliente, a.glsCliente, c.idNivel, d.glsNivel, e.idNivel as idFamilia, e.glsNivel as glsFamilia, a.idDocumento, '" & empresa & "', '" & sucursal & "', '" & Vendedor & "', '" & serie & "' " & _
                "from docventas a inner join docventasdet b " & _
                "on a.idDocVentas = b.idDocVentas and a.idSerie = b.idSerie and a.idDocumento = b.idDocumento and a.idEmpresa = b.idEmpresa and a.idSucursal = b.idSucursal " & _
                cempresas & _
                "inner join Productos c on b.idProducto = c.idProducto and b.idEmpresa = c.idEmpresa " & _
                "left join Niveles d on d.idNivel = c.idNivel and d.idEmpresa = a.idEmpresa left join Niveles e on d.idNivelPred = e.idNivel and d.idEmpresa = e.idEmpresa " & _
                "left join TiposdeCambio tc on year(tc.fecha) = year(a.FecEmision) and month(tc.fecha) = month(a.FecEmision) and day(tc.fecha) = day(a.FecEmision) " & _
                "where a.estDocventas in ('IMP','CAN') and a.idDocumento in('01','03','25') " & _
                "and a.FecEmision between cast('" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "' as date) and cast('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' as date) " & _
                "and a.idSucursal like '%" & Trim(txtCod_Sucursal.Text) & "%' and a.idPerVendedorCampo like '%" & Trim(txtCod_Vendedor.Text) & "%' " & _
                "and a.idSerie like '%" & Format(Trim(txt_serie.Text), "000") & "%' " & _
                "and a.idPerCliente like '%" & Trim(txtCod_Cliente.Text) & "%' and a.idEmpresa = '" & glsEmpresa & "' "
        Cn.Execute csql
        
        mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parReporte|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Report & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
        
    ElseIf Report = "4" Then
        csql = "insert into tempEstadistico(IDPRODUCTO, GLSPRODUCTO, CANTIDAD, CANTIDAD2, VVUNIT, PVUNIT, IDMONEDA, TIPOCAMBIO, IDPERCLIENTE, GLSCLIENTE, IDNIVEL, GLSNIVEL, IDFAMILIA, GLSFAMILIA, IDDOCUMENTO, EMPRESA, SUCURSAL, VENDEDOR, SERIE, diasVcto, idNivel02, GlsNivel02)" & _
                "Select b.idProducto, c.glsProducto, b.Cantidad, b.Cantidad2, " & _
                "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  IF(a.idMoneda = 'PEN', b.TotalVVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalVVNeto * tc.tcVenta)) " & _
                                                        "WHEN 'USD' THEN  IF(a.idMoneda = 'USD', b.TotalVVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalVVNeto / tc.tcVenta)) end," & _
                "CASE '" & Trim(txtCod_Moneda.Text) & "' WHEN 'PEN' THEN  IF(a.idMoneda = 'PEN', b.TotalPVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalPVNeto * tc.tcVenta)) " & _
                                                        "WHEN 'USD' THEN  IF(a.idMoneda = 'USD', b.TotalPVNeto, if(tc.tcVenta is null or tc.tcVenta = 0, 0, b.TotalPVNeto / tc.tcVenta)) end," & _
                "'" & Trim(txtCod_Moneda.Text) & "', isnull(tc.tcVenta) as TipoCambio, a.idPerCliente, a.glsCliente, c.idNivel, d.glsNivel, e.idNivel as idFamilia, e.glsNivel as glsFamilia, a.idDocumento, '" & empresa & "', '" & sucursal & "', '" & Vendedor & "', '" & serie & "', " & _
                "f.diasVcto, v.idNivel02, v.GlsNivel02 " & _
                "from docventas a inner join docventasdet b " & _
                "on a.idDocVentas = b.idDocVentas and a.idSerie = b.idSerie and a.idDocumento = b.idDocumento and a.idEmpresa = b.idEmpresa and a.idSucursal = b.idSucursal " & _
                cempresas & _
                "inner join Productos c on b.idProducto = c.idProducto and b.idEmpresa = c.idEmpresa " & _
                "inner join Clientes L on a.IdPerCliente = L.IdCliente and A.idEmpresa = L.idEmpresa " & _
                "left join Niveles d on d.idNivel = c.idNivel and d.idEmpresa = a.idEmpresa left join Niveles e on d.idNivelPred = e.idNivel and d.idEmpresa = e.idEmpresa " & _
                "inner join vw_niveles v on v.idNivel01 = d.idNivel and c.idEmpresa = v.idEmpresa " & _
                "inner join formaspagos f on L.idFormaPago = f.idFormaPago and L.idEmpresa = f.idEmpresa " & _
                "left join TiposdeCambio tc on year(tc.fecha) = year(a.FecEmision) and month(tc.fecha) = month(a.FecEmision) and day(tc.fecha) = day(a.FecEmision) " & _
                "where a.estDocventas in ('IMP','CAN') and a.idDocumento in('01','03','25') " & _
                "and a.FecEmision between cast('" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "' as date) and cast('" & Format(dtpFFinal.Value, "yyyy-mm-dd") & "' as date) " & _
                "and a.idSucursal like '%" & Trim(txtCod_Sucursal.Text) & "%' and a.idPerVendedorCampo like '%" & Trim(txtCod_Vendedor.Text) & "%' " & _
                "and a.idSerie like '%" & Format(Trim(txt_serie.Text), "000") & "%' " & _
                "and a.idPerCliente like '%" & Trim(txtCod_Cliente.Text) & "%' and a.idEmpresa = '" & glsEmpresa & "'"
        Cn.Execute csql
        
        mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMoneda|parReporte|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Moneda.Text) & "|" & Report & "|" & Format(dtpfInicio.Value, "yyyy-mm-dd") & "|" & Format(dtpFFinal.Value, "yyyy-mm-dd"), GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    Exit Sub
End Sub
