VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmRptVentasxResponsable 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas Por Responsable"
   ClientHeight    =   5220
   ClientLeft      =   4950
   ClientTop       =   1710
   ClientWidth     =   7020
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   5220
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4425
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   6765
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   5175
         TabIndex        =   9
         Top             =   4095
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1365
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
         Left            =   6195
         Picture         =   "frmRptVentasxResponsable.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1925
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
         Left            =   6195
         Picture         =   "frmRptVentasxResponsable.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1520
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
         Left            =   6195
         Picture         =   "frmRptVentasxResponsable.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1160
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   225
         Width           =   6400
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            Top             =   300
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   1
            Top             =   300
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
            Format          =   107085825
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   18
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3960
            TabIndex        =   17
            Top             =   375
            Width           =   420
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   14
         Left            =   135
         TabIndex        =   15
         Top             =   2340
         Width           =   6400
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   3960
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptGeneral 
            Caption         =   "General"
            Height          =   240
            Left            =   1305
            TabIndex        =   5
            Top             =   360
            Width           =   2025
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1185
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1140
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
         Container       =   "frmRptVentasxResponsable.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2120
         TabIndex        =   21
         Top             =   1140
         Width           =   4045
         _ExtentX        =   7144
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
         Container       =   "frmRptVentasxResponsable.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1185
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1515
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
         Container       =   "frmRptVentasxResponsable.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2120
         TabIndex        =   23
         Top             =   1515
         Width           =   4045
         _ExtentX        =   7144
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
         Container       =   "frmRptVentasxResponsable.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   1185
         TabIndex        =   4
         Tag             =   "TidPerCliente"
         Top             =   1920
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
         Container       =   "frmRptVentasxResponsable.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   2120
         TabIndex        =   25
         Tag             =   "TGlsCliente"
         Top             =   1920
         Width           =   4045
         _ExtentX        =   7144
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
         Container       =   "frmRptVentasxResponsable.frx":0B2A
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Frame FraOrdenDetallado 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   800
         Left            =   135
         TabIndex        =   20
         Top             =   3240
         Width           =   6400
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Fecha Emisión"
            Height          =   240
            Index           =   0
            Left            =   1395
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Documento"
            Height          =   240
            Index           =   1
            Left            =   4050
            TabIndex        =   8
            Top             =   360
            Width           =   2025
         End
      End
      Begin VB.Frame FraOrdenRes 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   800
         Left            =   135
         TabIndex        =   19
         Top             =   3240
         Visible         =   0   'False
         Width           =   6400
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   1
            Left            =   4005
            TabIndex        =   13
            Top             =   360
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Vendedor"
            Height          =   240
            Index           =   0
            Left            =   1350
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   26
         Top             =   1980
         Width           =   720
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   24
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   22
         Top             =   1185
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4635
      Width           =   1200
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4635
      Width           =   1200
   End
End
Attribute VB_Name = "frmRptVentasxResponsable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumerosFrames As String
Public GlsReporte As String
Public GlsForm As String
Public IndAgrupado

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
    
End Sub

Private Sub cmbAyudaSucursal_Click()

    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaVendedor_Click()
    
    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim fIni    As String, Ffin As String
Dim strMoneda As String
Dim StrMsgError As String
Dim COrden          As String
Dim CGlsReporte     As String
Dim GlsForm         As String
                    
    Screen.MousePointer = 11
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
    
    If OptDetallado.Value Then
        CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorVendedor_MonedaOriginal.rpt", "rptVentasPorVendedor.rpt")
        COrden = "GlsVendedor," & IIf(OptOrdenDet(0).Value, "Cast(FecEmision AS DateTime),AbreDocumento,idSerie,idDocVentas", "AbreDocumento,idSerie,idDocVentas,Cast(FecEmision AS DateTime)")
        GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
        mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta|ParTipo|ParOficial|parOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|1" & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & COrden, GlsForm, StrMsgError
    Else
        CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorVendedorGeneral_Moneda_Original.rpt", "rptVentasPorVendedorGeneral.rpt")
        COrden = IIf(OptOrdenRes(0).Value, "GlsVendedor", "TotalValorVenta desc")
        GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Responsable", " - Ordenado por Valor Venta")
        mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta|ParTipo|ParOficial|parOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|1" & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & COrden, GlsForm, StrMsgError
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
    
    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtCod_Moneda.Text = ""
    txtGls_Moneda.Text = "Moneda Original"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)

End Sub

Private Sub OptDetallado_Click()
    
    FraOrdenDetallado.Visible = True
    FraOrdenRes.Visible = False

End Sub

Private Sub OptGeneral_Click()
    
    FraOrdenDetallado.Visible = False
    FraOrdenRes.Visible = True

End Sub

Private Sub txtCod_Moneda_Change()
    
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    Else
        txtGls_Moneda.Text = "Moneda Original"
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

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text <> "" Then
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    Else
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    End If
    
End Sub
