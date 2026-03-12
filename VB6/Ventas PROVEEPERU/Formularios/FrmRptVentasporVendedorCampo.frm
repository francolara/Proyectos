VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRptVentasporVendedorCampo 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Vendedor de Campo"
   ClientHeight    =   5490
   ClientLeft      =   3645
   ClientTop       =   2250
   ClientWidth     =   7530
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7530
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   90
      TabIndex        =   17
      Top             =   45
      Width           =   7350
      Begin VB.Frame FraComision 
         Height          =   555
         Left            =   2340
         TabIndex        =   35
         Top             =   4095
         Visible         =   0   'False
         Width           =   2880
         Begin VB.OptionButton OptJefe 
            Caption         =   "Cobranza"
            Height          =   195
            Left            =   1620
            TabIndex        =   37
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton OptTodos 
            Caption         =   "Facturado"
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   225
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VB.CheckBox ChkComision 
         Caption         =   "Comision"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1260
         TabIndex        =   34
         Top             =   4305
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin VB.CommandButton cmbAyudaVendedor 
         Height          =   315
         Left            =   6710
         Picture         =   "FrmRptVentasporVendedorCampo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1960
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaMoneda 
         Height          =   315
         Left            =   6710
         Picture         =   "FrmRptVentasporVendedorCampo.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1610
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6710
         Picture         =   "FrmRptVentasporVendedorCampo.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1235
         Width           =   390
      End
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Documentos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5760
         TabIndex        =   14
         Top             =   4350
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Width           =   6915
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            Top             =   270
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
            Format          =   121896961
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   121896961
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   21
            Top             =   330
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3960
            TabIndex        =   20
            Top             =   330
            Width           =   420
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   14
         Left            =   180
         TabIndex        =   18
         Top             =   2430
         Width           =   6915
         Begin VB.OptionButton OptGeneralxvendedor 
            Caption         =   "General Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   765
            TabIndex        =   5
            Top             =   360
            Width           =   1755
         End
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3015
            TabIndex        =   6
            Top             =   360
            Width           =   1170
         End
         Begin VB.OptionButton OptGeneral 
            Caption         =   "General"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5040
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1230
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
         Container       =   "FrmRptVentasporVendedorCampo.frx":0A9E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2145
         TabIndex        =   26
         Top             =   1230
         Width           =   4545
         _ExtentX        =   8017
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
         Container       =   "FrmRptVentasporVendedorCampo.frx":0ABA
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Moneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Tag             =   "TidMoneda"
         Top             =   1600
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
         Container       =   "FrmRptVentasporVendedorCampo.frx":0AD6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Moneda 
         Height          =   315
         Left            =   2145
         TabIndex        =   29
         Top             =   1605
         Width           =   4545
         _ExtentX        =   8017
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
         Container       =   "FrmRptVentasporVendedorCampo.frx":0AF2
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Tag             =   "TidPerCliente"
         Top             =   1970
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
         Container       =   "FrmRptVentasporVendedorCampo.frx":0B0E
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   2145
         TabIndex        =   32
         Tag             =   "TGlsCliente"
         Top             =   1965
         Width           =   4545
         _ExtentX        =   8017
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
         Container       =   "FrmRptVentasporVendedorCampo.frx":0B2A
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Frame FraOrdenDetallado 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   180
         TabIndex        =   23
         Top             =   3330
         Width           =   6915
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4275
            TabIndex        =   9
            Top             =   315
            Width           =   1620
         End
         Begin VB.OptionButton OptOrdenDet 
            Caption         =   "Fecha Emisión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1215
            TabIndex        =   8
            Top             =   315
            Value           =   -1  'True
            Width           =   2025
         End
      End
      Begin VB.Frame FraOrdenRes 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   180
         TabIndex        =   24
         Top             =   3330
         Visible         =   0   'False
         Width           =   6915
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1080
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   1710
         End
         Begin VB.OptionButton OptOrdenRes 
            Caption         =   "Precio Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4140
            TabIndex        =   11
            Top             =   360
            Width           =   1800
         End
      End
      Begin VB.Frame FraGeneral 
         Appearance      =   0  'Flat
         Caption         =   " Orden "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   180
         TabIndex        =   22
         Top             =   3330
         Width           =   6915
         Begin VB.OptionButton OptordGeneral 
            Caption         =   "Ciente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4275
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton OptordGeneral 
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1215
            TabIndex        =   12
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.Label lblvendedor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
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
         Left            =   220
         TabIndex        =   33
         Top             =   2045
         Width           =   720
      End
      Begin VB.Label lbl_Moneda 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
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
         Left            =   225
         TabIndex        =   30
         Top             =   1650
         Width           =   570
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
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
         Left            =   225
         TabIndex        =   27
         Top             =   1275
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4890
      Width           =   1140
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4890
      Width           =   1140
   End
End
Attribute VB_Name = "FrmRptVentasporVendedorCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkComision_Click()
    If chkComision.Visible Then
        FraComision.Visible = True
    Else
        FraComision.Visible = False
    End If
End Sub

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
Dim fIni            As String, Ffin As String
Dim strMoneda       As String
Dim StrMsgError     As String
Dim COrden          As String
Dim CGlsReporte     As String
Dim GlsForm         As String
Dim strRUC          As String
    
    Screen.MousePointer = 11
    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strMoneda = IIf(Trim(txtCod_Moneda.Text) = "", "PEN", txtCod_Moneda.Text)

    
    If chkComision.Value Then
        If OptTodos.Value Then
            mostrarReporte "rptComisionPorVendedorCampoApimas.rpt", "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
        End If
        If OptJefe.Value Then
            mostrarReporte "rptComisionPorVendedorCampoApimasCobranza.rpt", "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
            'mostrarReporte "rptComisionPorVendedorCampoApimasXJefe.rpt", "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin, GlsForm, StrMsgError
        End If
    Else
        If OptGeneral.Value Then
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorVendedorCampo_Moneda_Original.rpt", "rptVentasPorVendedorCampo.rpt")
            COrden = IIf(OptordGeneral(0).Value, "GlsVendedorCampo", "GlsCliente desc")
            GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptordGeneral(0).Value, " - Ordenado por Vendedro de Campo", " - Ordenado por Cliente")
            mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta|ParTipo|ParOficial|parOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|0" & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & COrden, GlsForm, StrMsgError
            Exit Sub
        End If
         
        If OptDetallado.Value Then
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorVendedorCampoDetallado_Moneda_Original.rpt", "rptVentasPorVendedorCampoDetallado.rpt")
            COrden = "GlsVendedorCampo," & IIf(OptOrdenDet(0).Value, "Cast(FecEmision AS DateTime),AbreDocumento,idSerie,idDocVentas", "AbreDocumento,idSerie,idDocVentas,Cast(FecEmision AS DateTime)")
            GlsForm = "Reporte " & Me.Caption & " - Detallado" & IIf(OptOrdenDet(0).Value, " - Ordenado por Fecha de Emisión", " - Ordenado por Nº de Documento")
            mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta|ParTipo|ParOficial|parOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|0" & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & COrden, GlsForm, StrMsgError
        
        Else
            CGlsReporte = IIf(Trim(txtCod_Moneda.Text) = "", "rptVentasPorVendedorCampoGeneral_Moneda_Original.rpt", "rptVentasPorVendedorCampoGeneral.rpt")
            COrden = IIf(OptOrdenRes(0).Value, "GlsVendedorCampo", "TotalValorVenta desc")
            GlsForm = "Reporte " & Me.Caption & " - General" & IIf(OptOrdenRes(0).Value, " - Ordenado por Vendedor de Campo", " - Ordenado por Valor Venta")
            mostrarReporte CGlsReporte, "parEmpresa|parSucursal|parVendedor|parMoneda|parFecDesde|parFecHasta|ParTipo|ParOficial|parOrden", glsEmpresa & "|" & Trim(txtCod_Sucursal.Text) & "|" & txtCod_Vendedor.Text & "|" & strMoneda & "|" & fIni & "|" & Ffin & "|0" & "|" & IIf(ChkOficial.Visible, ChkOficial.Value, "1") & "|" & COrden, GlsForm, StrMsgError
        End If
    End If
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
Dim strRUC As String
    Me.top = 0
    Me.left = 0
    
    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"
    txtGls_Moneda.Text = "MONEDA ORIGINAL"
    dtpfInicio.Value = Format(Date, "dd/mm/yyyy")
    dtpFFinal.Value = Format(Date, "dd/mm/yyyy")
    ChkOficial.Visible = IIf(GlsVisualiza_Filtro_Documento = "S", True, False)
    
    If traerCampo("PARAMETROS", "VALPARAMETRO", "GLSPARAMETRO", "VISUALIZA_CLIENTESXVENDEDOR", True) = "S" Then
        If traerCampo("usuarios", "indJefe", "idUsuario", glsUser, True) = "0" Then
            txtCod_Vendedor.Text = Trim("" & glsUser)
            
            lblvendedor.Enabled = False
            txtCod_Vendedor.Enabled = False
            txtGls_Vendedor.Enabled = False
            cmbAyudaVendedor.Enabled = False
            
        Else
            lblvendedor.Enabled = True
            txtCod_Vendedor.Enabled = True
            txtGls_Vendedor.Enabled = True
            cmbAyudaVendedor.Enabled = True
        End If
    Else
        lblvendedor.Enabled = True
        txtCod_Vendedor.Enabled = True
        txtGls_Vendedor.Enabled = True
        cmbAyudaVendedor.Enabled = True
    End If
    strRUC = traerCampo("Empresas", "Ruc", "idEmpresa", glsEmpresa, False)
    
    If (strRUC = "20305948277" Or strRUC = "20987898989" Or strRUC = "20544632192") Then
        chkComision.Value = 0
        chkComision.Visible = True
    Else
        chkComision.Value = 0
        chkComision.Visible = False
    End If
    
End Sub

Private Sub OptDetallado_Click()
    
    FraOrdenDetallado.Visible = True
    FraOrdenRes.Visible = False
    fraGeneral.Visible = False

End Sub

Private Sub OptGeneral_Click()
    
    FraOrdenDetallado.Visible = False
    FraOrdenRes.Visible = False
    fraGeneral.Visible = True
    
End Sub

Private Sub OptGeneralxvendedor_Click()

    FraOrdenDetallado.Visible = False
    fraGeneral.Visible = False
    FraOrdenRes.Visible = True
    
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

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text <> "" Then
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    Else
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    End If
    
End Sub
