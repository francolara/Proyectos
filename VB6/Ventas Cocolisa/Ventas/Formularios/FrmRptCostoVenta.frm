VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRptCostoVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costo Venta"
   ClientHeight    =   3105
   ClientLeft      =   1965
   ClientTop       =   1980
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7335
   Begin VB.Frame Frame1 
      Height          =   2640
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   7290
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   105
         Width           =   7005
         Begin VB.CommandButton cmbAyudaSucursal 
            Height          =   315
            Left            =   6210
            Picture         =   "FrmRptCostoVenta.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   300
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Sucursal 
            Height          =   315
            Left            =   1335
            TabIndex        =   15
            Tag             =   "TidMoneda"
            Top             =   315
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
            Container       =   "FrmRptCostoVenta.frx":038A
            Estilo          =   1
            Vacio           =   -1  'True
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Sucursal 
            Height          =   315
            Left            =   2310
            TabIndex        =   16
            Top             =   315
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   556
            BackColor       =   12648447
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
            Container       =   "FrmRptCostoVenta.frx":03A6
            Vacio           =   -1  'True
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   765
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Rango de Fechas "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   1725
         Width           =   7005
         Begin MSComCtl2.DTPicker dtpfInicio 
            Height          =   315
            Left            =   1290
            TabIndex        =   9
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56950785
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtpFFinal 
            Height          =   315
            Left            =   4470
            TabIndex        =   10
            Top             =   315
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56950785
            CurrentDate     =   38667
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   12
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   705
            TabIndex        =   11
            Top             =   375
            Width           =   465
         End
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   5
         Left            =   135
         TabIndex        =   3
         Top             =   930
         Width           =   7005
         Begin VB.CommandButton cmbAyudaMoneda 
            Height          =   315
            Left            =   6210
            Picture         =   "FrmRptCostoVenta.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   270
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Moneda 
            Height          =   315
            Left            =   1350
            TabIndex        =   5
            Tag             =   "TidMoneda"
            Top             =   300
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
            Container       =   "FrmRptCostoVenta.frx":074C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Moneda 
            Height          =   315
            Left            =   2325
            TabIndex        =   6
            Top             =   300
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   556
            BackColor       =   12648447
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
            Container       =   "FrmRptCostoVenta.frx":0768
            Vacio           =   -1  'True
         End
         Begin VB.Label lbl_Moneda 
            Appearance      =   0  'Flat
            Caption         =   "Moneda:"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   7
            Top             =   375
            Width           =   765
         End
      End
   End
   Begin VB.CommandButton Btn_Salir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton Btn_Aceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1185
   End
End
Attribute VB_Name = "FrmRptCostoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Aceptar_Click()
Dim strFecIni, strFecFin, strmoneda, strSucursal As String
Dim strMsgError As String
On Error GoTo ERR
    strFecIni = Format(dtpfInicio.Value, "yyyy-mm-dd")
    strFecFin = Format(dtpFFinal.Value, "yyyy-mm-dd")
    strmoneda = Trim(txtCod_Moneda.Text)
    strSucursal = Trim(txtCod_Sucursal.Text)
    
    mostrarReporte "rptCostoVenta.rpt", "parEmpresa|parSucursal|parMoneda|parFechaIni|parFechaFin", glsEmpresa & "|" & strSucursal & "|" & strmoneda & "|" & strFecIni & "|" & strFecFin, GlsForm, strMsgError
    If strMsgError <> "" Then GoTo ERR
    Exit Sub
ERR:
    If strMsgError = "" Then strMsgError = ERR.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub Btn_Salir_Click()
   Unload Me
End Sub

Private Sub cmbAyudaMoneda_Click()
    mostrarAyuda "MONEDA", txtCod_Moneda, txtGls_Moneda
End Sub

Private Sub cmbAyudaSucursal_Click()
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
End Sub

Private Sub Form_Load()
    txtCod_Sucursal.Text = glsSucursal

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

 
