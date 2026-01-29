VERSION 5.00
Begin VB.Form frmMuestraImagen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Imagen del producto"
   ClientHeight    =   8310
   ClientLeft      =   3225
   ClientTop       =   2175
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   9495
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
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
      Height          =   480
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   7725
      Left            =   0
      ScaleHeight     =   511
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   629
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMuestraImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MostrarForm(ByVal strRuta As String)
    Picture1.Picture = LoadPicture(strRuta)
'    Picture1.ScaleMode = 3
'    Picture1.AutoRedraw = True
'    Picture1.PaintPicture Picture1.Picture, _
'        0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'    Picture1.AutoRedraw = False
    Me.Show 1
End Sub

Private Sub Command1_Click()
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then Unload Me

End Sub

Private Sub Picture1_Resize()
'    Picture1.AutoRedraw = True
'    Picture1.PaintPicture Picture1.Picture, _
'        0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'    Picture1.AutoRedraw = False
End Sub
