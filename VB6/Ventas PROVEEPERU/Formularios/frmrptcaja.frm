VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmrptcaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos en Caja"
   ClientHeight    =   2325
   ClientLeft      =   4245
   ClientTop       =   3510
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtnSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3375
      TabIndex        =   4
      Top             =   1755
      Width           =   1230
   End
   Begin VB.CommandButton Btnaceptar 
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
      Height          =   420
      Left            =   2070
      TabIndex        =   3
      Top             =   1755
      Width           =   1230
   End
   Begin VB.Frame fraReportes 
      Appearance      =   0  'Flat
      Caption         =   " Documento "
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
      Height          =   1575
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   45
      Width           =   6420
      Begin VB.TextBox txtnumdoc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   1080
         Width           =   1140
      End
      Begin VB.TextBox txtserie 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   720
         Width           =   645
      End
      Begin VB.CommandButton cmbAyudaTipoDoc 
         Height          =   315
         Left            =   5850
         Picture         =   "frmrptcaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   390
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1575
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   360
         Width           =   645
         _ExtentX        =   1138
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
         Container       =   "frmrptcaja.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   2250
         TabIndex        =   7
         Top             =   360
         Width           =   3565
         _ExtentX        =   6297
         _ExtentY        =   556
         BackColor       =   16777152
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
         Container       =   "frmrptcaja.frx":03A6
         Vacio           =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   10
         Top             =   405
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   9
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   8
         Top             =   1170
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmrptcaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GlsReporte As String
Public GlsForm As String

Private Sub Btnaceptar_Click()

    If Len(txtCod_Documento.Text) = 0 Or Len(txtCod_Documento.Text) > 2 Then
        MsgBox "Tipo de Documento Incorrecto,Verifque", vbInformation, App.Title
        txtCod_Documento.SetFocus
        Exit Sub
    End If

    If Len(txtserie.Text) = 0 Or Len(txtserie.Text) > 3 Then
         MsgBox "Numero de Serie Incorrecto,Verifque", vbInformation, App.Title
         txtserie.SetFocus
         Exit Sub
    End If

    If Len(txtnumdoc.Text) = 0 Or Len(txtnumdoc.Text) > 8 Then
         MsgBox "Numero de Documento Incorrecto,Verifque", vbInformation, App.Title
         txtnumdoc.SetFocus
        Exit Sub
    End If
    procesar

End Sub

Private Sub procesar()
On Error GoTo Err
Dim StrMsgError As String
Dim fec_inicio As String
Dim fec_fin    As String
    
    mostrarReporte "rptdocumento_caja.rpt", "parEmpresa|pardocu|parserie|parnumero", glsEmpresa & "|" & txtCod_Documento.Text & "|" & txtserie.Text & "|" & txtnumdoc.Text, "Ventas Documento en Caja", StrMsgError
    If StrMsgError <> "" Then GoTo Err
        
    Exit Sub
    
Err:
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub btnSalir_Click()
    
    Unload Me

End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0

End Sub

Private Sub txtnumdoc_LostFocus()
    
    txtnumdoc.Text = Format("" & txtnumdoc.Text, "00000000")

End Sub

Private Sub txtserie_LostFocus()
    
    txtserie.Text = Format("" & txtserie.Text, "000")

End Sub
