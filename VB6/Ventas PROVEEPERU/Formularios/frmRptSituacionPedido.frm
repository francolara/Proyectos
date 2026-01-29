VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmRptSituacionPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situación del Pedido"
   ClientHeight    =   3765
   ClientLeft      =   6570
   ClientTop       =   5280
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7065
   Begin VB.Frame Frame1 
      Height          =   3045
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmbAyudaProducto 
         Height          =   315
         Left            =   6195
         Picture         =   "frmRptSituacionPedido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1800
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6195
         Picture         =   "frmRptSituacionPedido.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2205
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   6195
         Picture         =   "frmRptSituacionPedido.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2565
         Width           =   390
      End
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
         Left            =   6180
         Picture         =   "frmRptSituacionPedido.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   330
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   855
         Width           =   6420
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1605
            TabIndex        =   1
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   131661825
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4605
            TabIndex        =   2
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   131661825
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4005
            TabIndex        =   13
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   990
            TabIndex        =   12
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Documento - Pedido "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   3525
         Visible         =   0   'False
         Width           =   6420
         Begin VB.TextBox txt_numdoc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4545
            TabIndex        =   4
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox txt_serie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   3
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            Height          =   210
            Left            =   855
            TabIndex        =   10
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   210
            Left            =   3645
            TabIndex        =   9
            Top             =   315
            Width           =   555
         End
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1095
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   330
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
         Container       =   "frmRptSituacionPedido.frx":0E28
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2055
         TabIndex        =   15
         Top             =   330
         Width           =   4095
         _ExtentX        =   7223
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
         Container       =   "frmRptSituacionPedido.frx":0E44
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Tag             =   "TidMoneda"
         Top             =   1800
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
         Container       =   "frmRptSituacionPedido.frx":0E60
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   2055
         TabIndex        =   21
         Top             =   1815
         Width           =   4095
         _ExtentX        =   7223
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
         Container       =   "frmRptSituacionPedido.frx":0E7C
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Tag             =   "TidMoneda"
         Top             =   2220
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
         Container       =   "frmRptSituacionPedido.frx":0E98
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2055
         TabIndex        =   23
         Top             =   2220
         Width           =   4095
         _ExtentX        =   7223
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
         Container       =   "frmRptSituacionPedido.frx":0EB4
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Tag             =   "TidMoneda"
         Top             =   2595
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
         Container       =   "frmRptSituacionPedido.frx":0ED0
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2055
         TabIndex        =   25
         Top             =   2595
         Width           =   4095
         _ExtentX        =   7223
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
         Container       =   "frmRptSituacionPedido.frx":0EEC
         Vacio           =   -1  'True
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   28
         Top             =   1890
         Width           =   645
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   27
         Top             =   2265
         Width           =   645
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   210
         TabIndex        =   26
         Top             =   2670
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   16
         Top             =   375
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   525
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3150
      Width           =   1740
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3150
      Width           =   1740
   End
End
Attribute VB_Name = "frmRptSituacionPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaCliente_Click()
    
    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
End Sub

Private Sub cmbAyudaProducto_Click()
    mostrarAyuda "PRODUCTOS", txtCod_Producto, txtGls_Producto
End Sub

Private Sub cmbAyudaSucursal_Click()
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
End Sub


Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim Fini                    As String, Ffin As String
Dim strMoneda               As String
Dim StrMsgError             As String
Dim COrden                  As String
Dim CGlsReporte             As String
Dim GlsForm                 As String
Dim X                       As Integer
Dim cGrupo                  As String
    
    
    Screen.MousePointer = 11
    Fini = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
    
    CGlsReporte = "rptSituacionPedido.rpt"
    COrden = ""
    'GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
    cGrupo = ""

    mostrarReporte CGlsReporte, "varEmpresa|varSucursal|varMoneda|varFechaIni|varFechaFin|varProducto|varOficial|varNiveles|varGrupo|varOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & strMoneda & "|" & Fini & "|" & Ffin & "|" & Trim(txtCod_Producto.Text) & "|0||" & cGrupo & "|" & COrden, Me.Caption, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
Exit Sub
Err:
    Screen.MousePointer = 0
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtCod_Moneda.Text = "PEN"
    
End Sub

Private Sub txt_numdoc_LostFocus()

    txt_NumDoc.Text = Format("" & txt_NumDoc.Text, "00000000")
    
End Sub

Private Sub txtCod_Cliente_Change()

    If txtCod_Cliente.Text <> "" Then
        txtGls_Cliente.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Cliente.Text, False)
    Else
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If
    
End Sub

Private Sub txt_Serie_LostFocus()

    txt_Serie.Text = Format("" & txt_Serie.Text, "000")
    
End Sub

Private Sub txtCod_Moneda_Change()
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    Else
        txtGls_Moneda.Text = "MONEDA ORIGINAL"
    End If
End Sub

Private Sub txtCod_Producto_Change()
    If txtCod_Producto.Text <> "" Then
        txtGls_Producto.Text = traerCampo("productos", "GlsProducto", "idProducto", txtCod_Producto.Text, True)
    Else
        txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    End If
End Sub

Private Sub txtCod_Sucursal_Change()
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    Me.Caption = Me.Caption & " - " & txtGls_Sucursal.Text
End Sub

