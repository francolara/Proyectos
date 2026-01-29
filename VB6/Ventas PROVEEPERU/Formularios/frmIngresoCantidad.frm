VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmIngresoCantidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingrese cantidad de etiquetas a imprimir"
   ClientHeight    =   2055
   ClientLeft      =   5865
   ClientTop       =   2490
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin CATControls.CATTextBox txtVal_Cantidad 
         Height          =   285
         Left            =   1950
         TabIndex        =   0
         Top             =   540
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
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
         Alignment       =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   -2147483640
         Container       =   "frmIngresoCantidad.frx":0000
         Estilo          =   3
         EnterTab        =   -1  'True
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   900
         TabIndex        =   5
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.Frame fraBotones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   90
      TabIndex        =   1
      Top             =   1260
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   2370
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmIngresoCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strCodProducto As String

Public Sub MostrarForm(ByVal strVarCodProducto As String, ByRef StrMsgError As String)

    strCodProducto = strVarCodProducto
    
    txtVal_Cantidad.Text = 0
    
    Me.Show 1
    
End Sub

Private Sub Command1_Click()
Dim StrMsgError As String

On Error GoTo Err

    If strCodProducto <> "" And txtVal_Cantidad.Value > 0 Then
        ImprimeCodigoBarra 2, strCodProducto, "", StrMsgError, txtVal_Cantidad.Value
        If StrMsgError <> "" Then GoTo Err
        Unload Me
    End If
    
Exit Sub
Err:
If StrMsgError = "" Then StrMsgError = Err.Description
MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub txtVal_Cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub
