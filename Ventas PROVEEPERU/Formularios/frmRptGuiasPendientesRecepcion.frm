VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmRptGuiasPendientesRecepcion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Guías Pendientes de Recepción"
   ClientHeight    =   1800
   ClientLeft      =   2475
   ClientTop       =   2715
   ClientWidth     =   6945
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
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1215
      Width           =   1200
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1215
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   6810
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6120
         Picture         =   "frmRptGuiasPendientesRecepcion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   375
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   360
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
         Container       =   "frmRptGuiasPendientesRecepcion.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   1995
         TabIndex        =   5
         Top             =   360
         Width           =   4095
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
         Container       =   "frmRptGuiasPendientesRecepcion.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   405
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmRptGuiasPendientesRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim StrMsgError As String

    Me.top = 0
    Me.left = 0
    txtGls_Sucursal.Text = "TODAS LAS SUCURSALES"

End Sub


Private Sub txtCod_Sucursal_Change()

    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If
    
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

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError     As String
Dim GlsReporte      As String
Dim GlsForm         As String
Dim strCodSucursal  As String

    Screen.MousePointer = 11
    strCodSucursal = Trim(txtCod_Sucursal.Text)
    
    GlsReporte = "rptRecepMercaderiaPendiente.rpt"
    GlsForm = Me.Caption
    
    mostrarReporte GlsReporte, "parEmpresa|parSucursal", glsEmpresa & "|" & strCodSucursal, GlsForm, StrMsgError
    
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
