VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmrptLiquidaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte - Liquidaciones"
   ClientHeight    =   4485
   ClientLeft      =   4935
   ClientTop       =   2865
   ClientWidth     =   6795
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
   ScaleHeight     =   4485
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3945
      Width           =   1200
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3945
      Width           =   1200
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
      Height          =   3795
      Left            =   45
      TabIndex        =   9
      Top             =   45
      Width           =   6675
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   2745
         Width           =   6375
         Begin VB.OptionButton OptResumido 
            Caption         =   "Resumido"
            Height          =   240
            Left            =   1125
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1860
         End
         Begin VB.OptionButton OptDetallado 
            Caption         =   "Detallado"
            Height          =   240
            Left            =   4095
            TabIndex        =   6
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   990
         Width           =   6375
         Begin VB.CommandButton CmdAyudaCamal 
            Height          =   315
            Left            =   5850
            Picture         =   "FrmrptLiquidaciones.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   315
            Width           =   390
         End
         Begin CATControls.CATTextBox txt_CodCamal 
            Height          =   315
            Left            =   810
            TabIndex        =   2
            Tag             =   "TidCamal"
            Top             =   315
            Width           =   1185
            _ExtentX        =   2090
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
            Container       =   "FrmrptLiquidaciones.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Camal 
            Height          =   315
            Left            =   2025
            TabIndex        =   16
            Top             =   315
            Width           =   3810
            _ExtentX        =   6720
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
            Container       =   "FrmrptLiquidaciones.frx":03A6
            Vacio           =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Camal"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   180
            TabIndex        =   17
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1890
         Width           =   6375
         Begin VB.OptionButton OptLiquidado 
            Caption         =   "Liquidado"
            Height          =   240
            Left            =   4095
            TabIndex        =   4
            Top             =   360
            Width           =   1320
         End
         Begin VB.OptionButton OptPendienteLiquidar 
            Caption         =   "Pendiente de Liquidar"
            Height          =   240
            Left            =   1125
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1860
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   810
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   135
         Width           =   6375
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
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
            Format          =   107544577
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
            Format          =   107544577
            CurrentDate     =   38667
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   900
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   375
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "FrmrptLiquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim fIni As String
Dim Ffin As String

    fIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    Ffin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    
    If OptPendienteLiquidar.Value = True Then
        If OptResumido.Value = True Then
            mostrarReporte "RptLiquidacionesPendientesResumido.Rpt", "ParEmpresa|ParSucursal|ParFecInicio|ParFecFinal|ParOpcion|ParCamal", glsEmpresa & "|" & glsSucursal & "|" & fIni & "|" & Ffin & "|" & "0" & "|" & txt_CodCamal.Text, "Pendientes por Liquidar", StrMsgError
        Else
            mostrarReporte "RptLiquidacionesPendientes.Rpt", "ParEmpresa|ParSucursal|ParFecInicio|ParFecFinal|ParOpcion|ParCamal", glsEmpresa & "|" & glsSucursal & "|" & fIni & "|" & Ffin & "|" & "0" & "|" & txt_CodCamal.Text, "Pendientes por Liquidar", StrMsgError
        End If
    Else
        If OptResumido.Value = True Then
            mostrarReporte "RptLiquidacionesCanceladasResumido.Rpt", "ParEmpresa|ParSucursal|ParFecInicio|ParFecFinal|ParOpcion|ParCamal", glsEmpresa & "|" & glsSucursal & "|" & fIni & "|" & Ffin & "|" & "1" & "|" & txt_CodCamal.Text, "Liquidadas", StrMsgError
        Else
            mostrarReporte "RptLiquidacionesCanceladas.Rpt", "ParEmpresa|ParSucursal|ParFecInicio|ParFecFinal|ParOpcion|ParCamal", glsEmpresa & "|" & glsSucursal & "|" & fIni & "|" & Ffin & "|" & "1" & "|" & txt_CodCamal.Text, "Liquidadas", StrMsgError
        End If
    End If
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub CmdAyudaCamal_Click()
    
    mostrarAyuda "UNIDADPRODUC", txt_CodCamal, txtGls_Camal
    If txt_CodCamal.Text <> "" Then SendKeys "{tab}"

End Sub

Private Sub cmdsalir_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.top = 0
    Me.left = 0
    OptPendienteLiquidar.Value = True
    OptResumido.Value = True
    
    dtpfInicio.Value = Format(getFechaSistema, "DD/MM/YYYY")
    dtpFFinal.Value = Format(getFechaSistema, "DD/MM/YYYY")
    
End Sub

Private Sub txt_CodCamal_Change()
    
    If txt_CodCamal.Text <> "" Then
        txtGls_Camal.Text = traerCampo("unidadproduccion", "DescUnidad", "CodUnidProd", txt_CodCamal.Text, False)
    Else
        txtGls_Camal.Text = "TODOS LOS CAMALES"
    End If

End Sub
