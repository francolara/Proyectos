VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form frmConfDocumentos 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Documentos"
   ClientHeight    =   7695
   ClientLeft      =   2070
   ClientTop       =   1635
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   Begin DXDBGRIDLibCtl.dxDBGrid gCabecera 
      Height          =   7530
      Left            =   75
      OleObjectBlob   =   "frmConfDocumentos.frx":0000
      TabIndex        =   0
      Top             =   60
      Width           =   13005
   End
End
Attribute VB_Name = "frmConfDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError As String
Dim csql As String

    Me.top = 0
    Me.left = 0
    
    ConfGrid gCabecera, True, False, False, False
    csql = "Select * From documentos"
    
    With gCabecera
        .DefaultFields = False
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CommandText = csql
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Open
        .Dataset.Active = True
        .KeyField = "idDocumento"
    End With
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub
