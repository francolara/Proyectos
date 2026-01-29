VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form FrmClientesPorVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Clientes por Vendedor"
   ClientHeight    =   1620
   ClientLeft      =   3750
   ClientTop       =   2460
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   7050
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
      Height          =   390
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton cmsalir 
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
      Height          =   390
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   855
      Index           =   15
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   6870
      Begin VB.CommandButton cmbAyudaVendedor 
         Height          =   315
         Left            =   6345
         Picture         =   "FrmClientesPorVendedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   320
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Vendedor 
         Height          =   315
         Left            =   1035
         TabIndex        =   0
         Tag             =   "TidPerCliente"
         Top             =   315
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
         Container       =   "FrmClientesPorVendedor.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Vendedor 
         Height          =   315
         Left            =   1965
         TabIndex        =   5
         Tag             =   "TGlsCliente"
         Top             =   315
         Width           =   4365
         _ExtentX        =   7699
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
         Container       =   "FrmClientesPorVendedor.frx":03A6
         Estilo          =   1
         Vacio           =   -1  'True
      End
      Begin VB.Label Label12 
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
         Left            =   195
         TabIndex        =   6
         Top             =   375
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmClientesPorVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaVendedor_Click()

    mostrarAyuda "VENDEDOR", txtCod_Vendedor, txtGls_Vendedor

End Sub

Private Sub cmdaceptar_Click()
On Error GoTo Err
Dim StrMsgError     As String

    mostrarReporte "rptClientesXVendedor.rpt", "parEmpresa|parVendedor", glsEmpresa & "|" & Trim(txtCod_Vendedor.Text), "Clientes por Vendedor", StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub cmsalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"

End Sub

Private Sub txtCod_Vendedor_Change()
    
    If txtCod_Vendedor.Text = "" Then
        txtGls_Vendedor.Text = "TODOS LOS VENDEDORES"
    Else
        txtGls_Vendedor.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Vendedor.Text, False)
    End If

End Sub
