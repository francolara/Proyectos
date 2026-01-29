VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmReporteAutorizaciones 
   Caption         =   "Reporte de Autorizaciones"
   ClientHeight    =   3975
   ClientLeft      =   4365
   ClientTop       =   2100
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   6870
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
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3375
      Width           =   1185
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3375
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   3120
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   6675
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6045
         Picture         =   "FrmReporteAutorizaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   780
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Tipo de Autorización "
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
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   2115
         Width           =   6240
         Begin VB.OptionButton opttipo 
            Caption         =   "Anulaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   945
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Modificación de Precios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3645
            TabIndex        =   5
            Top             =   360
            Width           =   2310
         End
      End
      Begin VB.CommandButton cmbAyudaUsuario 
         Height          =   315
         Left            =   6045
         Picture         =   "FrmReporteAutorizaciones.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   325
         Width           =   390
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
         TabIndex        =   9
         Top             =   1260
         Width           =   6240
         Begin MSComCtl2.DTPicker dtpdesde 
            Height          =   315
            Left            =   1515
            TabIndex        =   2
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   107610113
            CurrentDate     =   38667
         End
         Begin MSComCtl2.DTPicker dtphasta 
            Height          =   315
            Left            =   4515
            TabIndex        =   3
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
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
            Format          =   107610113
            CurrentDate     =   38667
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
            Left            =   3870
            TabIndex        =   11
            Top             =   375
            Width           =   420
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
            Left            =   810
            TabIndex        =   10
            Top             =   375
            Width           =   465
         End
      End
      Begin CATControls.CATTextBox txtCod_Usuario 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   330
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
         Container       =   "FrmReporteAutorizaciones.frx":0714
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Usuario 
         Height          =   315
         Left            =   1890
         TabIndex        =   13
         Top             =   330
         Width           =   4140
         _ExtentX        =   7303
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
         Container       =   "FrmReporteAutorizaciones.frx":0730
         Vacio           =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Tag             =   "TidMoneda"
         Top             =   780
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
         Container       =   "FrmReporteAutorizaciones.frx":074C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   1890
         TabIndex        =   17
         Top             =   780
         Width           =   4140
         _ExtentX        =   7303
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
         Container       =   "FrmReporteAutorizaciones.frx":0768
         Vacio           =   -1  'True
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
         TabIndex        =   18
         Top             =   825
         Width           =   645
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
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
         TabIndex        =   14
         Top             =   375
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmReporteAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
    
End Sub

Private Sub cmbAyudaUsuario_Click()
    
    mostrarAyuda "USUARIO", txtCod_Usuario, txtGls_Usuario
    
End Sub

Private Sub cmdaceptar_Click()

    PROCESA

End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    dtpdesde.Value = Format(Date, "DD/MM/YYYY")
    dtphasta.Value = Format(Date, "DD/MM/YYYY")

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    Me.Caption = Me.Caption & " - " & txtGls_Sucursal.Text
    
End Sub

Private Sub txtCod_Sucursal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 And KeyAscii <> 8 Then
        mostrarAyudaKeyascii KeyAscii, "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal
        KeyAscii = 0
    End If

End Sub

Private Sub txtCod_Usuario_KeyPress(KeyAscii As Integer)
 
    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "USUARIO", txtCod_Usuario, txtGls_Usuario
        KeyAscii = 0
        If txtCod_Usuario.Text <> "" Then SendKeys "{tab}"
    End If
 
End Sub

Private Sub PROCESA()
On Error GoTo Err
Dim rsReporte As ADODB.Recordset
Dim StrMsgError As String, strFecIni As String, strFecFin As String, strSucursal As String
    
    strFecIni = Format(dtpdesde.Value, "yyyy-mm-dd")
    strFecFin = Format(dtphasta.Value, "yyyy-mm-dd")
    strSucursal = Trim(txtCod_Sucursal.Text)
    
    If optTipo(0).Value = True Then
        Set rsReporte = DataProcedimiento("spu_ReporteAutorizacionesAnulaciones", StrMsgError, glsEmpresa, strSucursal, txtCod_Usuario.Text, strFecIni, strFecFin)
        If StrMsgError <> "" Then GoTo Err
        
        mostrarReporte rsReporte, "rptAutorizaAnulaciones.rpt", "Reporte de Autorizaciones", StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    If optTipo(1).Value = True Then
        Set rsReporte = DataProcedimiento("spu_ReporteAutorizacionesModificaPrecios", StrMsgError, glsEmpresa, strSucursal, txtCod_Usuario.Text, strFecIni, strFecFin)
        If StrMsgError <> "" Then GoTo Err
        
        mostrarReporte rsReporte, "rptAutorizaModificaPrecios.rpt", "Reporte de Autorizaciones", StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub
    
Err:
    If TypeName(rsReporte) = "Recordset" Then
        If rsReporte.State = 1 Then rsReporte.Close: Set rsReporte = Nothing
    End If
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub mostrarReporte(ByVal rsReporte As ADODB.Recordset, ByVal GlsReporte As String, ByVal GlsTitulo As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim vistaPrevia     As New frmReportePreview
Dim aplicacion      As New CRAXDRT.Application
Dim reporte         As CRAXDRT.Report
Dim gStrRutaRpts    As String

    Screen.MousePointer = 11
    gStrRutaRpts = App.Path + "\Reportes\"
    Set reporte = aplicacion.OpenReport(gStrRutaRpts & GlsReporte)
    If Not rsReporte.EOF And Not rsReporte.BOF Then
        reporte.Database.SetDataSource rsReporte, 3
        vistaPrevia.CRViewer91.ReportSource = reporte
        vistaPrevia.Caption = GlsTitulo
        vistaPrevia.CRViewer91.ViewReport
        vistaPrevia.CRViewer91.DisplayGroupTree = False
        Screen.MousePointer = 0
        vistaPrevia.WindowState = 2
    
        vistaPrevia.Show
    Else
        Screen.MousePointer = 0
        MsgBox "No existen Registros  Seleccionados", vbInformation, App.Title
    End If
    Screen.MousePointer = 0

    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    Exit Sub
    
Err:
    Screen.MousePointer = 0
    If rsReporte.State = 1 Then rsReporte.Close
    Set rsReporte = Nothing
    Set vistaPrevia = Nothing
    Set aplicacion = Nothing
    Set reporte = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

