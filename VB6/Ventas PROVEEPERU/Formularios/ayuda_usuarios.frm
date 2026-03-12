VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form ayuda_usuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Usuarios"
   ClientHeight    =   5895
   ClientLeft      =   4185
   ClientTop       =   2070
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   7485
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4620
         Left            =   90
         OleObjectBlob   =   "ayuda_usuarios.frx":0000
         TabIndex        =   1
         Top             =   225
         Width           =   7320
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   7485
      Begin VB.TextBox txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   945
         TabIndex        =   0
         Top             =   270
         Width           =   6405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   735
      End
   End
End
Attribute VB_Name = "ayuda_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw_limpia   As Boolean

Private Sub fill()
Dim csql    As String

    csql = "SELECT u.idusuario, p.glspersona " & _
           "FROM usuarios u INNER JOIN personas p " & _
           "ON u.idusuario = p.idpersona " & _
           "where idempresa = '" & glsEmpresa & "'"
    
    With dxDBGrid1
         .DefaultFields = False
         .Dataset.ADODataset.ConnectionString = strcn
         .Dataset.ADODataset.CursorLocation = clUseClient
         .Dataset.Active = False
         .Dataset.ADODataset.CommandText = csql
         .Dataset.DisableControls
         .Dataset.Active = True
         .KeyField = "idusuario"
    End With

End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    wusuario = "" & dxDBGrid1.Columns.ColumnByFieldName("idusuario").Value

    sw_limpia = True
    txtbusqueda.Text = ""
    sw_limpia = False
    
    Me.Hide
    
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
            wusuario = ""
            Me.Hide
    End Select
       
End Sub

Private Sub Form_Activate()
    
    dxDBGrid1.OptionEnabled = 0
    dxDBGrid1.Columns.FocusedIndex = 1
    dxDBGrid1.SetFocus
    dxDBGrid1.OptionEnabled = 1
    txtbusqueda_Change
    txtbusqueda.SetFocus
    dxDBGrid1.Dataset.First
    ConfGrid dxDBGrid1, False, False, False, False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub txtbusqueda_Change()
    
    If sw_limpia = False Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "idusuario LIKE '*" & txtbusqueda.Text & "*' OR " & " glspersona LIKE '*" & txtbusqueda.Text & "*'"
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
            dxDBGrid1.Dataset.Filter = "idusuario LIKE '*" & txtbusqueda.Text & "*' OR " & " glspersona LIKE '*" & txtbusqueda.Text & "*'"
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    dxDBGrid1.Dataset.First
    fill
    
End Sub
