VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   10725
   ClientTop       =   3810
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   4395
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   7080
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2460
         Left            =   150
         ScaleHeight     =   2460
         ScaleWidth      =   3300
         TabIndex        =   6
         Top             =   375
         Width           =   3300
      End
      Begin VB.Timer Timer1 
         Interval        =   150
         Left            =   4080
         Top             =   2280
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   6075
         Picture         =   "frmSplash.frx":000C
         Top             =   3450
         Width           =   960
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modulo de Ventas e Iventario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   3675
         TabIndex        =   7
         Top             =   375
         Width           =   3330
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4095
         TabIndex        =   4
         Top             =   2925
         Width           =   2910
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PowerSoft  :  S.A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   3150
         Width           =   2910
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmSplash.frx":08B5
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   105
         TabIndex        =   2
         Top             =   3540
         Width           =   5880
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Versión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5925
         TabIndex        =   5
         Top             =   2610
         Width           =   960
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Autorizado a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mAlpha As Long
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const RDW_INVALIDATE = &H1
Private Const RDW_ERASE = &H4
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_FRAME = &H400

Private Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Sub Fade(ByVal Enabled As Boolean)

    If Enabled = True Then
        Hide
        Timer1.Enabled = True
    Else
        Transparentar False, 0
    End If

End Sub

Private Sub Transparentar(ByVal Enabled As Boolean, ByVal Porcentaje As Long)

    If Enabled = True Then
        Dim tAlpha As Long
        tAlpha = Val(Porcentaje)
        If tAlpha < 1 Or tAlpha > 100 Then
            tAlpha = 70
        End If
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
    Else
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_LAYERED)
        Call RedrawWindow2(hwnd, 0&, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_FRAME Or RDW_ALLCHILDREN)
    End If

End Sub

Private Sub Form_Click()
    
    IniciarLogin

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then IniciarLogin

End Sub

Private Sub Form_Load()
Dim strRutaLogo As String

    strRutaLogo = App.Path & "\Logo\logoPW.jpg"
    
    If Len(Dir(strRutaLogo, vbArchive)) <> 0 Then
        Centrar_Imagen Picture1, strRutaLogo
    End If
    
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    App.Title = "SIAC"
    lblProductName.Caption = "Sistema de Ventas e intentario"
    
    Fade True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then IniciarLogin

End Sub

Private Sub Frame1_Click()
    
    IniciarLogin

End Sub

Private Sub Timer1_Timer()

    Show
    mAlpha = mAlpha + 10
    If mAlpha <= 100 Then
        Transparentar True, mAlpha
    ElseIf mAlpha = 200 Then
        Timer1.Enabled = False
        IniciarLogin
    End If

End Sub

Private Sub IniciarLogin()
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim StrMsgError As String

    Unload Me
    
    abrirConexion StrMsgError
    If StrMsgError <> "" Then GoTo Err
    frmLogin.Show 1

    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
