VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmCambiarClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3D3&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2355
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3D3&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2355
      Width           =   1140
   End
   Begin CATControls.CATTextBox txtUserName 
      Height          =   315
      Left            =   2625
      TabIndex        =   0
      Top             =   360
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      BackColor       =   16777215
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      ForeColor       =   -2147483640
      Container       =   "frmCambiarClave.frx":0000
   End
   Begin CATControls.CATTextBox txtPassword 
      Height          =   315
      Left            =   2625
      TabIndex        =   3
      Top             =   795
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      ForeColor       =   -2147483640
      PasswordChar    =   "X"
      Container       =   "frmCambiarClave.frx":001C
   End
   Begin CATControls.CATTextBox txtNuevoPassword 
      Height          =   315
      Left            =   2625
      TabIndex        =   6
      Top             =   1200
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      ForeColor       =   -2147483640
      PasswordChar    =   "X"
      Container       =   "frmCambiarClave.frx":0038
   End
   Begin CATControls.CATTextBox txtRepetirPassword 
      Height          =   315
      Left            =   2625
      TabIndex        =   8
      Top             =   1620
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      ForeColor       =   -2147483640
      PasswordChar    =   "X"
      Container       =   "frmCambiarClave.frx":0054
   End
   Begin VB.Label lblLabels 
      Caption         =   "Repetir Contraseña:"
      Height          =   270
      Index           =   3
      Left            =   660
      TabIndex        =   9
      Top             =   1650
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nueva Contraseña:"
      Height          =   270
      Index           =   2
      Left            =   660
      TabIndex        =   7
      Top             =   1230
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña Actual:"
      Height          =   270
      Index           =   1
      Left            =   660
      TabIndex        =   5
      Top             =   825
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   660
      TabIndex        =   4
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "frmCambiarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
    txtUserName.Text = traerCampo("usuarios", "varUsuario", "idUsuario", glsUser, True)
End Sub
