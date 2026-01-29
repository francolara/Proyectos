VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmVentasxCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas por Cliente"
   ClientHeight    =   5685
   ClientLeft      =   4065
   ClientTop       =   4215
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
   ScaleHeight     =   5685
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4965
      Left            =   60
      TabIndex        =   13
      Top             =   30
      Width           =   6765
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   1
         Left            =   6200
         Picture         =   "FrmVentasxCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2340
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaNivel 
         Height          =   315
         Index           =   0
         Left            =   6200
         Picture         =   "FrmVentasxCliente.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1950
         Width           =   390
      End
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         Height          =   240
         Left            =   8940
         TabIndex        =   10
         Top             =   3750
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox Chkzonas 
         Caption         =   "Zonas"
         Height          =   240
         Left            =   7890
         TabIndex        =   9
         Top             =   3900
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
         Left            =   6195
         Picture         =   "FrmVentasxCliente.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2745
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
         Picture         =   "FrmVentasxCliente.frx":0A9E
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
         Picture         =   "FrmVentasxCliente.frx":0E28
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
            Format          =   132382721
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
            Format          =   132382721
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
         Top             =   3090
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
         Container       =   "FrmVentasxCliente.frx":11B2
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
         Container       =   "FrmVentasxCliente.frx":11CE
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
         Container       =   "FrmVentasxCliente.frx":11EA
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
         Container       =   "FrmVentasxCliente.frx":1206
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Tag             =   "TidMoneda"
         Top             =   2715
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
         Container       =   "FrmVentasxCliente.frx":1222
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2145
         TabIndex        =   29
         Top             =   2715
         Width           =   4005
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
         Container       =   "FrmVentasxCliente.frx":123E
         Vacio           =   -1  'True
      End
      Begin VB.Frame FraOrdenRes 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         ForeColor       =   &H00000000&
         Height          =   810
         Left            =   7830
         TabIndex        =   21
         Top             =   2490
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
            Left            =   4800
            TabIndex        =   7
            Top             =   -180
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
         Top             =   3990
         Width           =   6330
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Fecha Emisión"
            Height          =   240
            Index           =   0
            Left            =   1320
            TabIndex        =   20
            Top             =   330
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Documento"
            Height          =   240
            Index           =   1
            Left            =   4020
            TabIndex        =   19
            Top             =   360
            Width           =   2025
         End
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   33
         Tag             =   "TidNivelPred"
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
         Container       =   "FrmVentasxCliente.frx":125A
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   315
         Index           =   0
         Left            =   2150
         TabIndex        =   34
         Top             =   1920
         Width           =   4005
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
         Container       =   "FrmVentasxCliente.frx":1276
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Nivel 
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   35
         Tag             =   "TidNivelPred"
         Top             =   2310
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
         Container       =   "FrmVentasxCliente.frx":1292
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Nivel 
         Height          =   315
         Index           =   1
         Left            =   2145
         TabIndex        =   36
         Top             =   2310
         Width           =   4005
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
         Container       =   "FrmVentasxCliente.frx":12AE
         Vacio           =   -1  'True
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sub Categoria"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   150
         TabIndex        =   38
         Top             =   2370
         Width           =   1020
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   150
         TabIndex        =   37
         Top             =   1950
         Width           =   690
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   150
         TabIndex        =   30
         Top             =   2760
         Width           =   570
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   150
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
         Left            =   150
         TabIndex        =   24
         Top             =   1185
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5130
      Width           =   1200
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
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

Private Sub cmbAyudaNivel_Click(Index As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    peso = Index + 1
    strCodTipoNivel = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
    strCondPred = ""
    If peso > 1 Then
        strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
    End If
    mostrarAyuda "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
    
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
Dim rsReporte       As New ADODB.Recordset
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report

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
            strSP = "spu_ListaVentasPorCliente"
            strTitulo = GlsForm
            strReporte = CGlsReporte
        Else
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorCliente_General_Moneda_Ori.rpt", "rptVentasPorCliente_General.rpt")
            COrden = IIf(OptOrdenRes(0).Value, "GlsCliente", "TotalValorVenta desc")
            GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Cliente", " - Ordenado por Precio Venta")
            strSP = GlsForm
            strTitulo = "Orden de Compra"
            strReporte = CGlsReporte
        End If
    End If
                 
    mostrarReporte CGlsReporte, "varEmpresa|varSucursal|varMoneda|varFecDesde|varFecHasta|varCliente|varOficial|varOrden|VarIdNivel|VarIdNivel2", glsEmpresa & "|" & txtCod_Sucursal.Text & "|" & CodMoneda & "|" & Fini & "|" & Ffin & "|" & Trim(txtCod_Cliente.Text) & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "0") & "|" & COrden & "|" & txtCod_Nivel(0).Text & "|" & txtCod_Nivel(1).Text, GlsForm, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
'    If strReporte = "" Then Screen.MousePointer = 0: Exit Sub
'    Set reporte = aplicacion.OpenReport(gStrRutaRpts & strReporte)
'    Set rsReporte = DataProcedimiento(strSP, StrMsgError, glsEmpresa, txtCod_Sucursal.Text, CodMoneda, Fini, Ffin, Trim(txtCod_Cliente.Text), IIf(ChkOficial.Visible, ChkOficial.Value, "0"), COrden)
'    If StrMsgError <> "" Then GoTo Err
'
'    If Not rsReporte.EOF And Not rsReporte.BOF Then
'         reporte.Database.SetDataSource rsReporte, 3
'         vistaPrevia.CRViewer91.ReportSource = reporte
'         vistaPrevia.Caption = strTitulo
'         vistaPrevia.CRViewer91.ViewReport
'         vistaPrevia.CRViewer91.DisplayGroupTree = False
'         Screen.MousePointer = 0
'         vistaPrevia.WindowState = 2
'         vistaPrevia.Show
'    Else
'        Screen.MousePointer = 0
'        MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
'    End If
'    Screen.MousePointer = 0
'    Set rsReporte = Nothing
'    Set vistaPrevia = Nothing
'    Set aplicacion = Nothing
'    Set reporte = Nothing
    
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
    
    txtGls_Nivel(0).Text = "TODOS"
    txtGls_Nivel(1).Text = "TODOS"
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

Private Sub txtCod_Nivel_Change(Index As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If indMovNivel Then Exit Sub
    txtGls_Nivel(Index).Text = traerCampo("niveles", "GlsNivel", "idNivel", txtCod_Nivel(Index).Text, True)
    
    indMovNivel = True
    For i = Index + 1 To txtCod_Nivel.Count - 1
        txtCod_Nivel(i).Text = ""
        txtGls_Nivel(i).Text = ""
    Next
    
    indMovNivel = False
    If glsNumNiveles = Index + 1 Then
'        If opt_Producto.Value Or opt_MateriaPrima.Value Then
'
'            If leeParametro("AYUDA_PRODUCTOS_CLIENTE") = "S" Then
'
'                CargarProductosCliente StrMsgError
'
'            Else
'
'                fill StrMsgError
'
'            End If
'
'        Else
'
'            fill StrMsgError
'
'        End If
        If StrMsgError <> "" Then GoTo Err
    End If
    If txtGls_Nivel(Index).Text = "" Then txtGls_Nivel(Index).Text = "TODOS"
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub txtCod_Nivel_KeyPress(Index As Integer, KeyAscii As Integer)
Dim peso As Integer
Dim strCodTipoNivel As String
Dim strCondPred As String
    
    If KeyAscii <> 13 Then
        peso = Index + 1
        strCodJerarquia = traerCampo("tiposniveles", "idTipoNivel", "peso", CStr(peso), True)
        strCondPred = ""
        If peso > 1 Then
            strCondPred = " AND idNivelPred = '" & txtCod_Nivel(Index - 1).Text & "'"
        End If
        'mostrarAyudaKeyascii KeyAscii, "NIVEL", txtCod_Nivel(Index), txtGls_Nivel(Index), " AND idTipoNivel = '" & strCodTipoNivel & "'" & strCondPred
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

