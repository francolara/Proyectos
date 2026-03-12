VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmConfDocVentas 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Documentos de Ventas"
   ClientHeight    =   9210
   ClientLeft      =   2055
   ClientTop       =   1560
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   9180
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   13200
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   140
         TabIndex        =   4
         Top             =   135
         Width           =   12900
         Begin VB.CommandButton cmbAyudaTipoDocumento 
            Height          =   315
            Left            =   6150
            Picture         =   "frmConfDocVentas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   225
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   1125
            TabIndex        =   0
            Top             =   225
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
            MaxLength       =   2
            Container       =   "frmConfDocVentas.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   2055
            TabIndex        =   6
            Top             =   225
            Width           =   4080
            _ExtentX        =   7197
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
            MaxLength       =   255
            Container       =   "frmConfDocVentas.frx":03A6
         End
         Begin CATControls.CATTextBox txt_Serie 
            Height          =   315
            Left            =   12105
            TabIndex        =   1
            Top             =   225
            Width           =   600
            _ExtentX        =   1058
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
            MaxLength       =   3
            Container       =   "frmConfDocVentas.frx":03C2
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   11535
            TabIndex        =   9
            Top             =   255
            Width           =   375
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Documento"
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
            Left            =   240
            TabIndex        =   7
            Top             =   260
            Width           =   810
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
         Height          =   2835
         Left            =   135
         OleObjectBlob   =   "frmConfDocVentas.frx":03DE
         TabIndex        =   3
         Top             =   930
         Width           =   12990
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
         Height          =   2910
         Left            =   135
         OleObjectBlob   =   "frmConfDocVentas.frx":4EE9
         TabIndex        =   8
         Top             =   3960
         Width           =   12990
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gTotales 
         Height          =   1980
         Left            =   135
         OleObjectBlob   =   "frmConfDocVentas.frx":9551
         TabIndex        =   10
         Top             =   7020
         Width           =   12990
      End
   End
End
Attribute VB_Name = "frmConfDocVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaTipoDocumento_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim csql As String
    
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento
    listar
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0
        
    ConfGrid gCabecera, True, False, False, False
    ConfGrid GDetalle, True, False, False, False
    ConfGrid gTotales, True, False, False, False
    
End Sub

Private Sub txt_Serie_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        listar
    End If

End Sub

Private Sub txtCod_Documento_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then
        mostrarAyudaKeyascii KeyAscii, "DOCUMENTOS", txtCod_Documento, txtGls_Documento
        KeyAscii = 0
        If txtCod_Documento.Text <> "" Then SendKeys "{tab}"
    End If

End Sub

Private Sub listar()

    If txtCod_Documento.Text = "" Then
        MsgBox "Ingrese Documento", vbInformation, App.Title
        Exit Sub
    End If
    
    If Trim(txt_serie.Text) = "" Then
        MsgBox "Ingrese Serie", vbInformation, App.Title
        Exit Sub
    End If

    csql = "Select * From objdocventas Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' AND idDocumento = '" & txtCod_Documento.Text & "'  AND idSerie = '" & txt_serie.Text & "'"
    With gCabecera
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "GlsObj"
    End With
    
    csql = "Select * From objdocventas Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' AND idDocumento = '" & txtCod_Documento.Text & "'  AND idSerie = '" & txt_serie.Text & "'"
    With GDetalle
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "GlsObj"
    End With
    
    csql = "Select * From objdocventas Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' AND idDocumento = '" & txtCod_Documento.Text & "'  AND idSerie = '" & txt_serie.Text & "'"
    With gTotales
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "GlsObj"
    End With
    
End Sub
