VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form ayuda_detraccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Códigos de  Detracción"
   ClientHeight    =   5970
   ClientLeft      =   2355
   ClientTop       =   1920
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5100
      Left            =   45
      TabIndex        =   4
      Top             =   810
      Width           =   7395
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4665
         Left            =   135
         OleObjectBlob   =   "ayuda_detraccion.frx":0000
         TabIndex        =   1
         Top             =   270
         Width           =   7155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   7395
      Begin VB.TextBox txtbusqueda 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   270
         Width           =   6360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
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
         Left            =   90
         TabIndex        =   3
         Top             =   315
         Width           =   735
      End
   End
End
Attribute VB_Name = "ayuda_detraccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw_limpia                       As Boolean
Dim CIdDetraccion                   As String

Public Sub MostrarForm(StrMsgError As String, PIdDetraccion As String)
On Error GoTo Err
    
    Me.Show 1
    
    PIdDetraccion = CIdDetraccion
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub fill()
On Error GoTo Err
Dim CSqlC                           As String
Dim StrMsgError                     As String

    CSqlC = "Select CODCONCEPTO,DESCRIPCION,PORCENTAJE FROM TB_CONCEP_DETRAC ORDER BY CODCONCEPTO"
    
    With dxDBGrid1
         .DefaultFields = False
         .Dataset.ADODataset.ConnectionString = strcn
         .Dataset.ADODataset.CursorLocation = clUseClient
         .Dataset.Active = False
         .Dataset.ADODataset.CommandText = CSqlC
         .Dataset.DisableControls
         .Dataset.Active = True
         .KeyField = "CODCONCEPTO"
    End With
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub dxDBGrid1_OnDblClick()
On Error GoTo Err
Dim StrMsgError                     As String

    CIdDetraccion = "" & dxDBGrid1.Columns.ColumnByFieldName("CODCONCEPTO").Value
     
    sw_limpia = True
    txtbusqueda.Text = ""
    sw_limpia = False
    
    Me.Hide
    
    Exit Sub
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
    MsgBox StrMsgError, vbInformation, App.Title
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    
    Select Case KeyCode
        Case 13:
            dxDBGrid1_OnDblClick
        Case 27:
            CIdDetraccion = ""
            Me.Hide
    End Select
       
End Sub

Private Sub Form_Activate()
    
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    dxDBGrid1.Columns.FocusedIndex = 1
    dxDBGrid1.SetFocus
    dxDBGrid1.OptionEnabled = 1
    txtbusqueda.SetFocus
    ConfGrid dxDBGrid1, False, False, False, False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub txtbusqueda_Change()
    
    If sw_limpia = False Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "CODCONCEPTO LIKE '*" & txtbusqueda.Text & "*' OR " & " DESCRIPCION LIKE '*" & txtbusqueda.Text & "*'"
        If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If

End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "CODCONCEPTO LIKE '*" & txtbusqueda.Text & "*' OR " & " DESCRIPCION LIKE '*" & txtbusqueda.Text & "*'"
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub

Private Sub Form_Load()

    fill
    
End Sub
