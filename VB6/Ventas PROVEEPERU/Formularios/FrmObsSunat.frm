VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form FrmObsSunat 
   Caption         =   "Observaciones SUNAT"
   ClientHeight    =   4725
   ClientLeft      =   4185
   ClientTop       =   3645
   ClientWidth     =   11190
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   11190
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11115
      Begin DXDBGRIDLibCtl.dxDBGrid GLista 
         Height          =   4395
         Left            =   90
         OleObjectBlob   =   "FrmObsSunat.frx":0000
         TabIndex        =   1
         Top             =   180
         Width           =   10935
      End
   End
End
Attribute VB_Name = "FrmObsSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CIdDocumento                    As String
Dim CIdSerie                        As String
Dim CIdDocVentas                    As String

Public Sub MostrarForm(StrMsgError As String, PIdDocumento As String, PIdSerie As String, PIdDocVentas As String)
On Error GoTo Err
    
    CIdDocumento = PIdDocumento
    CIdSerie = PIdSerie
    CIdDocVentas = PIdDocVentas
    
    Me.Show 1

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim StrMsgError                     As String
    
    ConfGrid gLista, True, True, False, False
    gLista.Options.Unset (egoCanInsert)
    
    Lista StrMsgError
    If StrMsgError <> "" Then GoTo Err
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Public Sub Lista(StrMsgError As String)
On Error GoTo Err
Dim CSqlC                           As String
Dim rsdatos                         As New ADODB.Recordset

    CSqlC = "Select 1 AS Item,'' AS FechaHora,'' AS NombreArchivo,'' AS ErrorSistema,A.MensajeEnvioSUNAT AS SunatDescription " & _
            "From DocVentas A " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdDocumento = '" & CIdDocumento & "' And A.IdSerie = '" & CIdSerie & "' " & _
            "And A.IdDocVentas = '" & CIdDocVentas & "' AND ISNULL(A.MensajeEnvioSUNAT,'') <> '' " & _
            ""
            
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open CSqlC, Cn, adOpenStatic, adLockOptimistic
        
    Set gLista.DataSource = rsdatos
    
'    With gLista
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = strcn
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = CSqlC
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "Item"
'    End With
'
    Me.Refresh

    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub
