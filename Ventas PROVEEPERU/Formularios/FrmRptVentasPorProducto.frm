VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRptVentasPorProducto 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas por Producto"
   ClientHeight    =   3075
   ClientLeft      =   6750
   ClientTop       =   5415
   ClientWidth     =   7035
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
   ScaleHeight     =   3075
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "COTIZACIONES"
      Height          =   2445
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   6855
      Begin VB.CheckBox ChkAgrupa 
         Caption         =   "Agrupado"
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   3330
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   5385
         TabIndex        =   9
         Top             =   3315
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   6250
         Picture         =   "FrmRptVentasPorProducto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1935
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6250
         Picture         =   "FrmRptVentasPorProducto.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1575
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaProducto 
         Height          =   315
         Left            =   6250
         Picture         =   "FrmRptVentasPorProducto.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1170
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   225
         Width           =   6400
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            Top             =   300
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   131989505
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   1
            Top             =   300
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   131989505
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   900
            TabIndex        =   18
            Top             =   330
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   17
            Top             =   330
            Width           =   420
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   800
         Index           =   14
         Left            =   180
         TabIndex        =   15
         Top             =   2430
         Visible         =   0   'False
         Width           =   6400
         Begin VB.OptionButton OptGeneral 
            Caption         =   "General"
            Height          =   240
            Left            =   1305
            TabIndex        =   5
            Top             =   360
            Width           =   2025
         End
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   3960
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
      End
      Begin CATControls.CATTextBox txtCod_Producto 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
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
         Container       =   "FrmRptVentasPorProducto.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Producto 
         Height          =   315
         Left            =   2115
         TabIndex        =   22
         Top             =   1185
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
         Container       =   "FrmRptVentasPorProducto.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1590
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
         Container       =   "FrmRptVentasPorProducto.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2115
         TabIndex        =   25
         Top             =   1590
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
         Container       =   "FrmRptVentasPorProducto.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   1950
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
         Container       =   "FrmRptVentasPorProducto.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2115
         TabIndex        =   28
         Top             =   1965
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
         Container       =   "FrmRptVentasPorProducto.frx":0B2A
         Vacio           =   -1  'True
      End
      Begin VB.Frame FraOrdenRes 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   800
         Left            =   180
         TabIndex        =   20
         Top             =   4605
         Visible         =   0   'False
         Width           =   6400
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Valor Venta"
            Height          =   240
            Index           =   1
            Left            =   4005
            TabIndex        =   8
            Top             =   360
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Producto"
            Height          =   240
            Index           =   0
            Left            =   1350
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
      End
      Begin VB.Frame FraOrdenDetallado 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   800
         Left            =   180
         TabIndex        =   19
         Top             =   4425
         Width           =   6400
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Fecha Emisión"
            Height          =   240
            Index           =   0
            Left            =   1350
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Documento"
            Height          =   240
            Index           =   1
            Left            =   4005
            TabIndex        =   13
            Top             =   360
            Width           =   2025
         End
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   29
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   26
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   270
         TabIndex        =   23
         Top             =   1260
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   1300
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2550
      Width           =   1300
   End
End
Attribute VB_Name = "FrmRptVentasPorProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaMoneda_Click()

    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

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
Dim cNiveles                As String
Dim X                       As Integer
Dim cGrupo                  As String
    
    OptDetallado.Value = True
    
    Screen.MousePointer = 11
    Fini = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
    
    For X = 1 To glsNumNiveles
        cNiveles = cNiveles & "idNivel" & Format(X, "00") & ", GlsNivel" & Format(X, "00") & ","
    Next X
    
    If OptDetallado.Value Then
        CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorProducto" & Format(X - 1, "00") & "_Moneda_Original.rpt", "rptVentasPorProducto" & Format(X - 1, "00") & ".rpt")
        COrden = "" '"GlsProducto,Documento," & IIf(OptOrdenDet(0).Value, "Cast(FecEmision AS DateTime),Documento", "Documento,Cast(FecEmision AS DateTime)")
        GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
        cGrupo = "" ' "Group By GlsProducto,Documento"
        cNiveles = ""
    Else
        If ChkAgrupa.Value = 1 Then
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorProductoGeneral" & Format(X - 1, "00") & "_Moneda_Original.rpt", "rptVentasPorProductoGeneral" & Format(X - 1, "00") & ".rpt")
            COrden = IIf(OptOrdenRes(0).Value, "GlsProducto", "TotalValorVenta desc")
            GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Producto", " - Ordenado por Valor Venta")
            cGrupo = "Group By GlsProducto"
        Else
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorProductoGeneral" & Format(X - 1, "00") & "_Moneda_Original_Sin_Familia.rpt", "rptVentasPorProductoGeneral" & Format(X - 1, "00") & "_Sin_Familia.rpt")
            COrden = IIf(OptOrdenRes(0).Value, "GlsProducto", "TotalValorVenta desc")
            GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Producto", " - Ordenado por Valor Venta")
            cGrupo = "Group By GlsProducto"
        End If
    End If
    
    mostrarReporte CGlsReporte, "varEmpresa|varSucursal|varMoneda|varFechaIni|varFechaFin|varProducto|varOficial|varNiveles|varGrupo|varOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & strMoneda & "|" & Fini & "|" & Ffin & "|" & Trim(txtCod_Producto.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & cNiveles & "|" & cGrupo & "|" & COrden, Me.Caption, StrMsgError
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
    
    Me.top = 0
    Me.left = 0
    
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Producto.Text = "TODOS LOS PRODUCTOS"
    txtCod_Moneda.Text = "PEN"
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)
    
    If OptGeneral.Value Then
        ChkAgrupa.Visible = True
    Else
        ChkAgrupa.Visible = False
    End If

End Sub

Private Sub OptGeneral_Click()
    
    If OptGeneral.Value Then
        ChkAgrupa.Visible = True
    Else
        ChkAgrupa.Visible = False
    End If
    
    FraOrdenDetallado.Visible = False
    'FraOrdenRes.Visible = True

End Sub

Private Sub OptDetallado_Click()
    
    ChkAgrupa.Visible = False
    
    FraOrdenDetallado.Visible = True
    'FraOrdenRes.Visible = False

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

Private Sub txtCod_Producto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 8 Then
        txtCod_Producto.Text = ""
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
