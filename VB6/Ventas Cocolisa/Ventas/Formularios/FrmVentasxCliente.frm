VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmVentasxCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas por Cliente"
   ClientHeight    =   5460
   ClientLeft      =   10080
   ClientTop       =   1635
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   90
      TabIndex        =   13
      Top             =   90
      Width           =   6765
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   5265
         TabIndex        =   10
         Top             =   4185
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox Chkzonas 
         Caption         =   "Zonas"
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   4185
         Visible         =   0   'False
         Width           =   825
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
         Left            =   6200
         Picture         =   "FrmVentasxCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1935
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
         Left            =   6200
         Picture         =   "FrmVentasxCliente.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   6200
         Picture         =   "FrmVentasxCliente.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1125
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   225
         Width           =   6330
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   121962497
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4515
            TabIndex        =   1
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   121962497
            CurrentDate     =   38667
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   16
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   14
         Left            =   180
         TabIndex        =   14
         Top             =   2385
         Width           =   6330
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   4005
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptGeneral 
            Caption         =   "General"
            Height          =   240
            Left            =   1350
            TabIndex        =   5
            Top             =   360
            Width           =   2025
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1200
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
         Container       =   "FrmVentasxCliente.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2150
         TabIndex        =   23
         Top             =   1140
         Width           =   4000
         _ExtentX        =   7064
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
         Container       =   "FrmVentasxCliente.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Cliente 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1545
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
         Container       =   "FrmVentasxCliente.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Cliente 
         Height          =   315
         Left            =   2150
         TabIndex        =   26
         Top             =   1545
         Width           =   4000
         _ExtentX        =   7064
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
         Container       =   "FrmVentasxCliente.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   1965
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
         Container       =   "FrmVentasxCliente.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2150
         TabIndex        =   29
         Top             =   1965
         Width           =   4000
         _ExtentX        =   7064
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
         Container       =   "FrmVentasxCliente.frx":0B2A
         Vacio           =   -1  'True
      End
      Begin VB.Frame FraOrdenRes 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   810
         Left            =   180
         TabIndex        =   21
         Top             =   3240
         Visible         =   0   'False
         Width           =   6330
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   1
            Left            =   4005
            TabIndex        =   8
            Top             =   360
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Cliente"
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
         Height          =   810
         Left            =   180
         TabIndex        =   18
         Top             =   3240
         Width           =   6330
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Fecha Emisión"
            Height          =   240
            Index           =   0
            Left            =   1395
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Documento"
            Height          =   240
            Index           =   1
            Left            =   4050
            TabIndex        =   19
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
         Left            =   225
         TabIndex        =   30
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   225
         TabIndex        =   27
         Top             =   1590
         Width           =   480
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   225
         TabIndex        =   24
         Top             =   1185
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4950
      Width           =   1200
   End
End
Attribute VB_Name = "FrmVentasxCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaCliente_Click()

    mostrarAyuda "CLIENTE", txtCod_Cliente, txtGls_Cliente

End Sub

Private Sub cmbAyudaMoneda_Click()
    
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
    If txtCod_Moneda.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim Fini As String
Dim Ffin As String
Dim CodMoneda As String
Dim COrden          As String
Dim CGlsReporte     As String
Dim GlsForm         As String

    Fini = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    CodMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)
    
    If Chkzonas.Value = 1 Then
        If OptDetallado.Value Then
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorClienteZonas02.rpt", "rptVentasPorClienteZonas.rpt")
            COrden = "GlsCliente," & IIf(OptOrdenDet(0).Value, "Cast(FecEmision AS Date),AbreDocumento,idSerie,idDocVentas", "AbreDocumento,idSerie,idDocVentas,Cast(FecEmision AS DateTime)")
            GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
        Else
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorCliente_General_Moneda_Ori_Zona.rpt", "rptVentasPorCliente_General_Zona.rpt")
            COrden = IIf(OptOrdenRes(0).Value, "GlsCliente", "TotalValorVenta desc")
            GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Cliente", " - Ordenado por Precio Venta")
        End If
    Else
        If OptDetallado.Value Then
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorCliente02.rpt", "rptVentasPorCliente.rpt")
            COrden = "GlsCliente," & IIf(OptOrdenDet(0).Value, "Cast(FecEmision AS Date),AbreDocumento,idSerie,idDocVentas", "AbreDocumento,idSerie,idDocVentas,Cast(FecEmision AS DateTime)")
            GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
        Else
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorCliente_General_Moneda_Ori.rpt", "rptVentasPorCliente_General.rpt")
            COrden = IIf(OptOrdenRes(0).Value, "GlsCliente", "TotalValorVenta desc")
            GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Cliente", " - Ordenado por Precio Venta")
        End If
    End If
    
    mostrarReporte CGlsReporte, "varEmpresa|varSucursal|varMoneda|varFecDesde|varFecHasta|varCliente|varOficial|varOrden", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & CodMoneda & "|" & Fini & "|" & Ffin & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "0") & "|" & COrden, GlsForm, StrMsgError
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
    
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    txtGls_Moneda.Text = "MONEDA ORIGINAL"
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
        txtGls_Cliente.Text = "TODOS LOS CLIENTES"
    End If
    
End Sub

Private Sub txtCod_Moneda_Change()
    
    If Len(Trim(txtCod_Moneda.Text)) > 0 Then
        txtGls_Moneda.Text = traerCampo("monedas", "GlsMoneda", "idMoneda", txtCod_Moneda.Text, False)
    Else
        txtGls_Moneda.Text = "MONEDA ORIGINAL"
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
