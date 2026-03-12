VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmRptVentasporVendedorCampo_Comisión 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisión por Vendedor de Campo"
   ClientHeight    =   5550
   ClientLeft      =   4095
   ClientTop       =   2745
   ClientWidth     =   7080
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7080
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4995
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4995
      Width           =   1185
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   " Tipo "
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   14
      Left            =   90
      TabIndex        =   21
      Top             =   3240
      Width           =   6915
      Begin VB.OptionButton OptGeneral 
         Caption         =   "General"
         Height          =   240
         Left            =   1575
         TabIndex        =   29
         Top             =   270
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   240
         Left            =   4230
         TabIndex        =   22
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   15
      Left            =   90
      TabIndex        =   15
      Top             =   2475
      Width           =   6915
      Begin VB.CommandButton cmbAyudaVendedor 
         Height          =   315
         Left            =   6120
         Picture         =   "FrmRptVentasporVendedorCampo_Comisión.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   195
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   285
         Left            =   1380
         TabIndex        =   17
         Tag             =   "TidPerCliente"
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
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
         MaxLength       =   8
         Container       =   "FrmRptVentasporVendedorCampo_Comisión.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   285
         Left            =   2370
         TabIndex        =   18
         Tag             =   "TGlsCliente"
         Top             =   225
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         Container       =   "FrmRptVentasporVendedorCampo_Comisión.frx":03A6
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   195
         TabIndex        =   19
         Top             =   285
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   5
      Left            =   90
      TabIndex        =   6
      Top             =   1665
      Width           =   6915
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   6060
         Picture         =   "FrmRptVentasporVendedorCampo_Comisión.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Tag             =   "TidMoneda"
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
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
         MaxLength       =   8
         Container       =   "FrmRptVentasporVendedorCampo_Comisión.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   285
         Left            =   2325
         TabIndex        =   9
         Top             =   300
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   503
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "FrmRptVentasporVendedorCampo_Comisión.frx":0768
         Vacio           =   -1  'True
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         Caption         =   "Moneda:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   225
         TabIndex        =   10
         Top             =   375
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   765
      Index           =   4
      Left            =   90
      TabIndex        =   5
      Top             =   855
      Width           =   6915
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   4905
         TabIndex        =   20
         Top             =   495
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6060
         Picture         =   "FrmRptVentasporVendedorCampo_Comisión.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   165
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1335
         TabIndex        =   12
         Tag             =   "TidMoneda"
         Top             =   180
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
         Container       =   "FrmRptVentasporVendedorCampo_Comisión.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2310
         TabIndex        =   13
         Top             =   180
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         BackColor       =   16777152
         Enabled         =   0   'False
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
         Container       =   "FrmRptVentasporVendedorCampo_Comisión.frx":0B2A
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Sucursal:"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   " Rango de Fechas "
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   6915
      Begin MSComCtl2.DTPicker dtpfInicio 
         Height          =   315
         Left            =   1515
         TabIndex        =   1
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107085825
         CurrentDate     =   38667
      End
      Begin MSComCtl2.DTPicker dtpFFinal 
         Height          =   315
         Left            =   4515
         TabIndex        =   2
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107085825
         CurrentDate     =   38667
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3960
         TabIndex        =   4
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   900
         TabIndex        =   3
         Top             =   375
         Width           =   465
      End
   End
   Begin VB.Frame FraOrdenDetallado 
      Appearance      =   0  'Flat
      Caption         =   " Orden "
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   90
      TabIndex        =   23
      Top             =   4050
      Width           =   6915
      Begin VB.OptionButton OptOrdenDet 
         Caption         =   "Fecha Emisión"
         Height          =   240
         Index           =   0
         Left            =   1530
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton OptOrdenDet 
         Caption         =   "Documento"
         Height          =   240
         Index           =   1
         Left            =   4185
         TabIndex        =   24
         Top             =   270
         Width           =   2025
      End
   End
   Begin VB.Frame FraGeneral 
      Appearance      =   0  'Flat
      Caption         =   " Orden "
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   90
      TabIndex        =   26
      Top             =   4050
      Visible         =   0   'False
      Width           =   6915
      Begin VB.OptionButton OptOrdenRes 
         Caption         =   "Precio Venta"
         Height          =   240
         Index           =   1
         Left            =   4185
         TabIndex        =   28
         Top             =   270
         Width           =   2025
      End
      Begin VB.OptionButton OptOrdenRes 
         Caption         =   "Vendedor"
         Height          =   240
         Index           =   0
         Left            =   1530
         TabIndex        =   27
         Top             =   270
         Value           =   -1  'True
         Width           =   2025
      End
   End
End
Attribute VB_Name = "FrmRptVentasporVendedorCampo_Comisión"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Command1_Click()
Dim fIni            As String, Ffin As String
Dim strMoneda       As String
Dim StrMsgError     As String
Dim COrden          As String
Dim CGlsReporte     As String
Dim GlsForm         As String
Screen.MousePointer = 11
On Error GoTo Err

    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
   strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
 
    If OptGeneral.Value Then
        CGlsReporte = "rptVentasPorVendedorCampoGeneralComision.rpt"
        COrden = IIf(OptOrdenRes(0).Value, "GlsVendedorCampo", "TotalValorVenta desc")
        GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Vendedor de Campo", " - Ordenado por Valor Venta")
    Else
        CGlsReporte = "rptVentasPorVendedorCampoComision.rpt"
        COrden = IIf(OptOrdenDet(0).Value, "Cast(FecEmision AS DateTime),Documento", "Cast(FecEmision AS DateTime),Documento")
        GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
    End If
         
    mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta|ParOficial|parOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & Trim(txtCod_Vendedor.Text) & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & COrden, GlsForm, StrMsgError
    If StrMsgError <> "" Then GoTo Err
  
Exit Sub
Err:
Screen.MousePointer = 0
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
'txtGls_Moneda.Text = "MONEDA ORIGINAL"
txtCod_Moneda.Text = "PEN"
dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)

End Sub
Private Sub OptDetallado_Click()
    FraOrdenDetallado.Visible = True
    FraGeneral.Visible = False
End Sub

Private Sub OptGeneral_Click()
   FraGeneral.Visible = True
   FraOrdenDetallado.Visible = False
End Sub
 
Private Sub txtCod_Moneda_Change()
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    'Else
        'txtGls_Moneda.Text = "MONEDA ORIGINAL"
    End If
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
'    mostrarAyudaKeyascii KeyAscii, "MONEDA", txtCod_Moneda, txtGls_Moneda
'    KeyAscii = 0
'    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"
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

