VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "catcontrols.ocx"
Begin VB.Form frmConfiguraImpresion 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configura Impresión de Documentos de Venta"
   ClientHeight    =   9180
   ClientLeft      =   2115
   ClientTop       =   1560
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListado 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   9135
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   11280
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   150
         TabIndex        =   7
         Top             =   150
         Width           =   10965
         Begin VB.CommandButton cmbAyudaTipoDocumento 
            Height          =   315
            Left            =   6420
            Picture         =   "frmConfiguraImpresion.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   225
            Width           =   390
         End
         Begin CATControls.CATTextBox txtCod_Documento 
            Height          =   315
            Left            =   1125
            TabIndex        =   2
            Top             =   225
            Width           =   690
            _ExtentX        =   1217
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
            Container       =   "frmConfiguraImpresion.frx":038A
            Estilo          =   1
            EnterTab        =   -1  'True
         End
         Begin CATControls.CATTextBox txtGls_Documento 
            Height          =   315
            Left            =   1830
            TabIndex        =   0
            Top             =   225
            Width           =   4575
            _ExtentX        =   8070
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
            Container       =   "frmConfiguraImpresion.frx":03A6
         End
         Begin CATControls.CATTextBox txt_Serie 
            Height          =   315
            Left            =   10110
            TabIndex        =   1
            Top             =   210
            Width           =   690
            _ExtentX        =   1217
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
            Container       =   "frmConfiguraImpresion.frx":03C2
            Estilo          =   1
            EnterTab        =   -1  'True
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
            Left            =   195
            TabIndex        =   10
            Top             =   270
            Width           =   810
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
            Left            =   9645
            TabIndex        =   9
            Top             =   255
            Width           =   375
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
         Height          =   2475
         Left            =   135
         OleObjectBlob   =   "frmConfiguraImpresion.frx":03DE
         TabIndex        =   3
         Top             =   1200
         Width           =   11040
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
         Height          =   2970
         Left            =   135
         OleObjectBlob   =   "frmConfiguraImpresion.frx":539C
         TabIndex        =   4
         Top             =   4020
         Width           =   11040
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gTotales 
         Height          =   1710
         Left            =   135
         OleObjectBlob   =   "frmConfiguraImpresion.frx":99F2
         TabIndex        =   5
         Top             =   7320
         Width           =   10995
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   7065
         Width           =   615
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   3765
         Width           =   555
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cabecera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   945
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmConfiguraImpresion"
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

    csql = "Select Identificador,GlsObs,indImprime,impX,impY,impLongitud,intNumFilas From objdocventas " & _
            "Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'C' AND idDocumento = '" & txtCod_Documento.Text & "' AND idSerie = '" & txt_serie.Text & "' and trim(GlsCampo) <> ''"
    With gCabecera
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
    csql = "Select Identificador,etiqueta,indImprime,impX,impY,impLongitud From objdocventas " & _
            "Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'D' AND idDocumento = '" & txtCod_Documento.Text & "'  AND idSerie = '" & txt_serie.Text & "' and trim(GlsCampo) <> ''"
    With GDetalle
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
    csql = "Select Identificador,GlsObs,indImprime,impX,impY,impLongitud From objdocventas Where idEmpresa = '" & glsEmpresa & "' AND tipoObj = 'T' AND idDocumento = '" & txtCod_Documento.Text & "' AND idSerie = '" & txt_serie.Text & "' and trim(GlsCampo) <> ''"
    With gTotales
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "Identificador"
    End With
    
End Sub
