VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F41D1D30-7878-4923-8CB3-6CCACDC9C9DE}#1.0#0"; "CATControls.ocx"
Begin VB.Form frmMotivosGuiaImp 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configura Impresión Motivos de Guías"
   ClientHeight    =   7815
   ClientLeft      =   4665
   ClientTop       =   1305
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   105
      TabIndex        =   3
      Top             =   15
      Width           =   8160
      Begin VB.CommandButton cmbAyudaTipoDocumento 
         Height          =   315
         Left            =   5760
         Picture         =   "frmMotivosGuiaImp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   280
         Width           =   390
      End
      Begin CATControls.CATTextBox txt_Serie 
         Height          =   315
         Left            =   7200
         TabIndex        =   1
         Top             =   285
         Width           =   780
         _ExtentX        =   1376
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
         Container       =   "frmMotivosGuiaImp.frx":038A
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtCod_Documento 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Tag             =   "TidDocumento"
         Top             =   285
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
         Container       =   "frmMotivosGuiaImp.frx":03A6
         Estilo          =   1
         EnterTab        =   -1  'True
      End
      Begin CATControls.CATTextBox txtGls_Documento 
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   285
         Width           =   3765
         _ExtentX        =   6641
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
         Container       =   "frmMotivosGuiaImp.frx":03C2
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   915
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
         Left            =   6720
         TabIndex        =   4
         Top             =   330
         Width           =   375
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
      Height          =   6885
      Left            =   105
      OleObjectBlob   =   "frmMotivosGuiaImp.frx":03DE
      TabIndex        =   2
      Top             =   840
      Width           =   8160
   End
End
Attribute VB_Name = "frmMotivosGuiaImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAyudaTipoDocumento_Click()
On Error GoTo Err
Dim StrMsgError As String
Dim csql As String
        
    mostrarAyuda "DOCUMENTOS", txtCod_Documento, txtGls_Documento
        
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub Form_Load()
    
    Me.top = 0
    Me.left = 0
    ConfGrid gCabecera, True, False, False, False

End Sub
 
Private Sub txt_Serie_KeyPress(KeyAscii As Integer)
On Error GoTo Err
Dim StrMsgError As String

    If KeyAscii = 13 Then
        listar StrMsgError
        If StrMsgError <> "" Then GoTo Err
    End If
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub listar(ByRef StrMsgError As String)
On Error GoTo Err
Dim rst As New ADODB.Recordset
Dim csql As String
Dim csqlInsert As String

    If txt_serie.Text = "" Then Exit Sub
    
    txt_serie.Text = Format(txt_serie.Text, "0000")
    csql = "Select impmotivostraslados.identificador,impmotivostraslados.idMotivoTraslado,m.GlsMotivoTraslado,impmotivostraslados.impX,impmotivostraslados.impY " & _
           "From impmotivostraslados " & _
           "Inner Join motivostraslados m " & _
           " On impmotivostraslados.idMotivoTraslado = m.idMotivoTraslado And impmotivostraslados.idDocumento=m.idDocumento " & _
           "Where impmotivostraslados.idEmpresa = '" & glsEmpresa & "' " & _
           "AND impmotivostraslados.idSerie = '" & txt_serie.Text & "' " & _
           "AND impmotivostraslados.idDocumento = '" & txtCod_Documento.Text & "' "
           
    rst.Open csql, Cn, adOpenForwardOnly, adLockReadOnly
    If rst.EOF Or rst.BOF Then
        If MsgBox("El numero de serie " & txt_serie.Text & " no existe, ¿Desea registrarla como una serie valida?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Cn.BeginTrans
            
            csqlInsert = "INSERT INTO impmotivostraslados (idEmpresa,idSerie,idMotivoTraslado,impX,impY,idDocumento) " & _
                        "SELECT '" & glsEmpresa & "','" & txt_serie.Text & "',m.idMotivoTraslado,0,0 ,'" & txtCod_Documento.Text & "' FROM motivostraslados m " & _
                        "WHERE m.idDocumento = '" & txtCod_Documento.Text & "' "
            Cn.Execute (csqlInsert)
            
            Cn.CommitTrans
            
        Else
            If rst.State = 1 Then rst.Close: Set rst = Nothing
            gCabecera.Dataset.Active = False
            Set gCabecera.DataSource = Nothing
            Exit Sub
        End If
    End If
    
    With gCabecera
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "identificador"
    End With
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    
    Exit Sub
    
Err:
    If rst.State = 1 Then rst.Close: Set rst = Nothing
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
