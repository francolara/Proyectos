VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmBusca_Entidad 
   Caption         =   "Busqueda "
   ClientHeight    =   1785
   ClientLeft      =   8745
   ClientTop       =   5490
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   7335
   Begin VB.Frame FraBuscaEntidad 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7260
      Begin VB.CommandButton BtnAceptar 
         Caption         =   "Aceptar"
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1125
         Width           =   1140
      End
      Begin CATControls.CATTextBox txt_RUC 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Tag             =   "Truc"
         Top             =   270
         Width           =   2250
         _ExtentX        =   3969
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
         MaxLength       =   11
         Container       =   "FrmBusca_Entidad.frx":0000
         Estilo          =   4
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_RazonSocial 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Tag             =   "TapellidoPaterno"
         Top             =   720
         Width           =   5970
         _ExtentX        =   10530
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
         MaxLength       =   80
         Container       =   "FrmBusca_Entidad.frx":001C
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin VB.Label lblRUC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
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
         Left            =   135
         TabIndex        =   5
         Top             =   345
         Width           =   450
      End
      Begin VB.Label lblRazonSocial 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Razón Social"
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
         Left            =   135
         TabIndex        =   3
         Top             =   765
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmBusca_Entidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCliRUC    As String
Dim strCliGlosa  As String

Private Sub BtnAceptar_Click()
Dim StrMsgError As String
On Error GoTo ERR

    If Trim(txt_RUC.Text) = "" Then
        StrMsgError = "Ingrese RUC"
        txt_RUC.SetFocus
        txt_RUC.SelStart = 0
        If StrMsgError <> "" Then GoTo ERR
    End If
    
    strCliRUC = txt_RUC.Text
    strCliGlosa = txtGls_RazonSocial.Text
    
    Unload Me

    Exit Sub
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
    
    strCliRUC = ""
    strCliGlosa = ""
    
End Sub

Public Sub MostrarForm(ByRef strRUC As String, ByRef strRazonSocial As String)
Dim StrMsgError As String
On Error GoTo ERR
    Me.Show 1
    
    strRUC = strCliRUC
    strRazonSocial = strCliGlosa
    
    Exit Sub
ERR:
    If StrMsgError = "" Then StrMsgError = ERR.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
