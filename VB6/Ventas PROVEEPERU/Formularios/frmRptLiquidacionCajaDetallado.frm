VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmRptLiquidacionCajaDetallado 
   Caption         =   "Liquidación de Caja - Detallado"
   ClientHeight    =   3015
   ClientLeft      =   3750
   ClientTop       =   2055
   ClientWidth     =   7035
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
   ScaleHeight     =   3015
   ScaleWidth      =   7035
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2475
      Width           =   1200
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2475
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   6855
      Begin VB.CommandButton cmbAyudaCaja 
         Height          =   315
         Left            =   6225
         Picture         =   "frmRptLiquidacionCajaDetallado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1665
         Width           =   390
      End
      Begin VB.CommandButton cmbAyudaUsuario 
         Height          =   315
         Left            =   6225
         Picture         =   "frmRptLiquidacionCajaDetallado.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1260
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Fecha "
         ForeColor       =   &H00000000&
         Height          =   765
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   6420
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   2940
            TabIndex        =   0
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Format          =   107544577
            CurrentDate     =   38667
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Del"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2340
            TabIndex        =   7
            Top             =   360
            Width           =   225
         End
      End
      Begin CATControls.CATTextBox txtCod_Usuario 
         Height          =   315
         Left            =   1275
         TabIndex        =   1
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
         Container       =   "frmRptLiquidacionCajaDetallado.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Usuario 
         Height          =   315
         Left            =   2250
         TabIndex        =   9
         Top             =   1275
         Width           =   3960
         _ExtentX        =   6985
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
         Container       =   "frmRptLiquidacionCajaDetallado.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Caja 
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Tag             =   "TidMoneda"
         Top             =   1680
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
         Container       =   "frmRptLiquidacionCajaDetallado.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Caja 
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         Top             =   1680
         Width           =   3960
         _ExtentX        =   6985
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
         Container       =   "frmRptLiquidacionCajaDetallado.frx":0768
         Vacio           =   -1  'True
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Caja"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   1725
         Width           =   315
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   1305
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmRptLiquidacionCajaDetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError     As String
Dim GlsReporte      As String
Dim GlsForm         As String
Dim strMovCaja      As String, csucursal As String, fCorte As String
Dim rsTemp          As New ADODB.Recordset

    GlsReporte = "rptLiquidacionCajaDet.rpt"
    GlsForm = Me.Caption
    fCorte = Format(DtpFecha.Value, "yyyy-mm-dd")

    csql = "SELECT m.idMovCaja,m.IdSucursal " & _
            "FROM movcajas m " & _
            "WHERE m.idEmpresa = '" & glsEmpresa & "' " & _
            "AND m.idUsuario = '" & txtCod_Usuario.Text & "' " & _
            "AND m.idCaja = '" & txtCod_Caja.Text & "' " & _
            "AND DATE_FORMAT(m.FecCaja ,'%d/%m/%Y') = DATE_FORMAT('" & fCorte & "','%d/%m/%Y')"
    rsTemp.Open csql, Cn, adOpenForwardOnly, adLockOptimistic
    If Not rsTemp.EOF Then
        strMovCaja = "" & rsTemp.Fields("idMovCaja")
        csucursal = "" & rsTemp.Fields("IdSucursal")
    Else
        StrMsgError = "No hay caja disponible para la fecha indicada"
        GoTo Err
    End If
    If rsTemp.State = 1 Then rsTemp.Close: Set rsTemp = Nothing
        
    If Trim("" & traerCampo("parametros", "ValParametro", "GlsParametro", "FORMATO_LIQUIDACION", True)) = "2" Then
        mostrarReporte "rptLiquidacionCajaDet_formato2.rpt", "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & csucursal & "|" & strMovCaja, GlsForm, StrMsgError
    Else
        mostrarReporte GlsReporte, "parEmpresa|parSucursal|parMovCaja", glsEmpresa & "|" & csucursal & "|" & strMovCaja, GlsForm, StrMsgError
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
Dim StrMsgError As String

    Me.top = 0
    Me.left = 0
    
    DtpFecha.Value = Format(Date, "dd/mm/yyyy")
    
End Sub


Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIO", txtCod_Usuario, txtGls_Usuario
        KeyAscii = 0
        If txtCod_Usuario.Text <> "" Then SendKeys "{tab}"
    End If
    
End Sub

Private Sub cmbAyudaUsuario_Click()
    
    mostrarAyuda "USUARIO", txtCod_Usuario, txtGls_Usuario

End Sub

Private Sub txtCod_Caja_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "CAJASUSUARIOFILTRO", txtCod_Caja, txtGls_Caja, "AND u.idUsuario = '" & txtCod_Usuario.Text & "'"
        KeyAscii = 0
        If txtCod_Caja.Text <> "" Then SendKeys "{tab}"
    End If
 
End Sub

Private Sub cmbAyudaCaja_Click()
    
    mostrarAyuda "CAJASUSUARIOFILTRO", txtCod_Caja, txtGls_Caja, "AND u.idUsuario = '" & txtCod_Usuario.Text & "'"

End Sub

