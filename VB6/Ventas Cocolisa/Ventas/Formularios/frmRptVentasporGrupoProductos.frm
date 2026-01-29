VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmRptVentasporGrupoProductos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas por Grupo de Productos"
   ClientHeight    =   3390
   ClientLeft      =   6375
   ClientTop       =   2715
   ClientWidth     =   6915
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
   ScaleHeight     =   3390
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   3525
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2835
      Width           =   1200
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2835
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   6720
      Begin VB.CheckBox ChkAgrupado 
         Caption         =   "Agrupado por Vendedor de Campo"
         Height          =   195
         Left            =   3555
         TabIndex        =   4
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2940
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   6120
         Picture         =   "frmRptVentasporGrupoProductos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1650
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6120
         Picture         =   "frmRptVentasporGrupoProductos.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1290
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   6330
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1380
            TabIndex        =   0
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   106758145
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4380
            TabIndex        =   1
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   106758145
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   765
            TabIndex        =   10
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3825
            TabIndex        =   9
            Top             =   375
            Width           =   420
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1050
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1275
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
         Container       =   "frmRptVentasporGrupoProductos.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2000
         TabIndex        =   12
         Top             =   1275
         Width           =   4100
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
         Container       =   "frmRptVentasporGrupoProductos.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1050
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1650
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
         Container       =   "frmRptVentasporGrupoProductos.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2000
         TabIndex        =   15
         Top             =   1650
         Width           =   4100
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
         Container       =   "frmRptVentasporGrupoProductos.frx":0768
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   1725
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   1320
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmRptVentasporGrupoProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim GlsReporte     As String
Dim GlsForm         As String
Dim fIni            As String, Ffin As String
Dim strMoneda       As String, strCodSucursal As String

    Screen.MousePointer = 11
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
    strCodSucursal = Trim(txtCod_Sucursal.Text)
    
    GlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorVendedor_MonedaOriginal.rpt", "rptVentasPorVendedor.rpt")
    GlsForm = Me.Caption
    
    If ChkAgrupado.Value = 1 Then
        mostrarReporte "rptVentasGrupoProductoVendedor.rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    Else
        mostrarReporte left(GlsReporte, Len(GlsReporte) - 4) & ".rpt", "parEmpresa|parSucursal|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & strCodSucursal & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Screen.MousePointer = 0
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
Dim StrMsgError As String
Dim strNumeros() As String
Dim intTop  As Integer
Dim intForm As Integer
Dim indTipoProd As Boolean
Dim i As Integer

    Me.top = 0
    Me.left = 0
    
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtCod_Moneda.Text = glsMonVentas
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    
End Sub

Private Sub txtCod_Sucursal_Change()

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
End Sub

Private Sub txtCod_Sucursal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Sucursal.Text = ""
    End If

End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Moneda_Change()
    
    txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)

End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
        KeyAscii = 0
        If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

End Sub

