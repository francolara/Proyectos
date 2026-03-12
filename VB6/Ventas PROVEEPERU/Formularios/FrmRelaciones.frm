VERSION 5.00
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form FrmRelaciones 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Relación de Documentos"
   ClientHeight    =   3105
   ClientLeft      =   2190
   ClientTop       =   1260
   ClientWidth     =   7365
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7365
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   7170
      Begin VB.CommandButton cmbAyudaSucursal 
         Height          =   315
         Left            =   6480
         Picture         =   "FrmRelaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   370
         Width           =   390
      End
      Begin VB.Frame fraReportes 
         Appearance      =   0  'Flat
         Caption         =   " Documento "
         ForeColor       =   &H00000000&
         Height          =   1260
         Index           =   11
         Left            =   225
         TabIndex        =   7
         Top             =   855
         Width           =   6735
         Begin VB.CommandButton cmbAyudaTipoDoc 
            Height          =   315
            Left            =   6210
            Picture         =   "FrmRelaciones.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   315
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   945
            TabIndex        =   1
            Tag             =   "TidMoneda"
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
            Container       =   "FrmRelaciones.frx":0714
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   1905
            TabIndex        =   9
            Top             =   315
            Width           =   4275
            _ExtentX        =   7541
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
            Container       =   "FrmRelaciones.frx":0730
            Vacio           =   -1  'True
         End
         Begin CATControls.CATTextBox txt_Serie 
            Height          =   315
            Left            =   945
            TabIndex        =   2
            Top             =   720
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
            MaxLength       =   4
            Container       =   "FrmRelaciones.frx":074C
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtNum_Documento 
            Height          =   315
            Left            =   3645
            TabIndex        =   3
            Top             =   720
            Width           =   2535
            _ExtentX        =   4471
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
            Container       =   "FrmRelaciones.frx":0768
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "T/D"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   195
            TabIndex        =   12
            Top             =   375
            Width           =   240
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Número"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   2970
            TabIndex        =   11
            Top             =   765
            Width           =   555
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   225
            TabIndex        =   10
            Top             =   765
            Width           =   375
         End
      End
      Begin CATControls.CATTextBox txtCod_Sucursal 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Tag             =   "TidMoneda"
         Top             =   375
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
         Container       =   "FrmRelaciones.frx":0784
         Estilo          =   1
         Vacio           =   -1  'True
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Sucursal 
         Height          =   315
         Left            =   2130
         TabIndex        =   14
         Top             =   375
         Width           =   4320
         _ExtentX        =   7620
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
         Container       =   "FrmRelaciones.frx":07A0
         Vacio           =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   315
         TabIndex        =   15
         Top             =   420
         Width           =   645
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2565
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2565
      Width           =   1230
   End
End
Attribute VB_Name = "FrmRelaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GlsReporte As String
Public GlsForm As String

Private Sub cmbAyudaSucursal_Click()
    
    mostrarAyuda "SUCURSAL", txtCod_Sucursal, txtGls_Sucursal

End Sub

Private Sub cmbAyudaTipoDoc_Click()
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento

End Sub

Private Sub Command1_Click()
On Error GoTo Err
Dim StrMsgError As String

    validaFormSQL Me, StrMsgError
    If StrMsgError <> "" Then GoTo Err
                
    GlsReporte = "rptRelacionDoc.rpt"
    GlsForm = "Reporte de Relacion de Documentos"
                
    Imprimer StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Command2_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"

End Sub

Private Sub Imprimer(ByRef StrMsgError As String)
On Error GoTo Err
Dim num         As String
Dim serie       As String
Dim doc         As String
Dim sucursal    As String
    
    sucursal = txtCod_Sucursal.Text
    doc = Format(Trim(txtCod_Documento.Text), "00")
    serie = Trim(txt_serie.Text)
    num = Format(Trim(txtNum_Documento.Text), "00000000")

    mostrarReporte GlsReporte, "parEmpresa|parSucursal|parDoc|parCod|parSerie", glsEmpresa & "|" & sucursal & "|" & doc & "|" & num & "|" & serie, GlsForm, StrMsgError
    If StrMsgError <> "" Then GoTo Err

    txt_serie.Text = serie
    txtNum_Documento.Text = num
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub txtCod_Documento_Change()
    
    txtGls_Documento.Text = traerCampo("Documentos", "GlsDocumento", "idDocumento", txtCod_Documento.Text, False)

End Sub

Private Sub txtCod_Sucursal_Change()
    
    If txtCod_Sucursal.Text <> "" Then
        txtGls_Sucursal.Text = traerCampo("personas", "GlsPersona", "idPersona", txtCod_Sucursal.Text, False)
    Else
        txtGls_Sucursal.Text = "TODOS LAS SUCURSALES"
    End If

End Sub
