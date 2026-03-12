VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form FrmAyudaLotes_Vales 
   Caption         =   "Lista de Tallas"
   ClientHeight    =   5295
   ClientLeft      =   8475
   ClientTop       =   3555
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7515
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
      Height          =   735
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   7395
      Begin VB.TextBox txtbusqueda 
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
         TabIndex        =   4
         Top             =   315
         Width           =   735
      End
   End
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
      Height          =   4470
      Left            =   75
      TabIndex        =   2
      Top             =   750
      Width           =   7395
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4095
         Left            =   120
         OleObjectBlob   =   "FrmAyudaLotes_Vales.frx":0000
         TabIndex        =   1
         Top             =   180
         Width           =   7155
      End
   End
End
Attribute VB_Name = "FrmAyudaLotes_Vales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw_limpia   As Boolean
Dim codigo As String
Dim Descripcion As String
Dim DescripcionLote As String
Dim strNewPeriodoinv    As String
Private Sub fill(codproducto As String, codalmacen As String, xindVale As String)
Dim csql     As String
Dim rsdatos  As New ADODB.Recordset

If xindVale = "I" Then
    csql = "select l.idlote,l.glslote,0 Stock from LOTES  l left join productosalmacenporlote pl " & _
            "on l.idlote = pl.idlote " & _
            "and l.idempresa = pl.idempresa " & _
            "and l.idsucursal = pl.idsucursal " & _
            "and l.estado = 'ACT' " & _
            "where L.IDSUCURSAL = '" & glsSucursal & "' " & _
            "and l.idempresa = '" & glsEmpresa & "' " & _
            "GROUP BY l.idlote,l.glslote"
            
            dxDBGrid1.Columns.ColumnByFieldName("Stock").Visible = False
Else
    csql = "Select B.IdLote,C.glslote,C.FechaLote,Sum(B.Cantidad * CASE WHEN B.TipoVale = 'I' THEN 1 ELSE -1 END) Stock " & _
            "From ValesCab A " & _
            "Inner Join ValesDetLotes B " & _
            "On A.IdEmpresa = B.IdEmpresa And A.IdSucursal = B.IdSucursal And A.TipoVale = B.TipoVale And A.IdValesCab = B.IdValesCab " & _
            "Inner Join Lotes C " & _
            "On B.IdEmpresa = C.IdEmpresa And B.IdLote = C.IdLote " & _
            "Where A.IdEmpresa = '" & glsEmpresa & "' And A.IdPeriodoInv = '" & glsCodPeriodoINV & "' " & _
            "And A.IdAlmacen = '" & codalmacen & "' And A.EstValeCab <> 'ANU' " & _
            "And B.IdProducto = '" & Trim("" & codproducto) & "' " & _
            "and C.estado = 'ACT' " & _
            "Group By B.IdLote,C.glslote,C.FechaLote " & _
            "Order By C.FechaLote"
            
            dxDBGrid1.Columns.ColumnByFieldName("Stock").Visible = True
End If
            '"And pl.idproducto = '" & codproducto & "' "
           '"AND pl.idalmacen = '" & codalmacen & "'
           
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Open csql, Cn, adOpenStatic, adLockOptimistic
        
    Set dxDBGrid1.DataSource = rsdatos

'    With dxDBGrid1
'         .DefaultFields = False
'         .Dataset.ADODataset.ConnectionString = strcn
'         .Dataset.ADODataset.CursorLocation = clUseClient
'         .Dataset.Active = False
'         .Dataset.ADODataset.CommandText = csql
'         .Dataset.DisableControls
'         .Dataset.Active = True
'         .KeyField = "idlote"
'    End With

End Sub

Private Sub dxDBGrid1_OnDblClick()

    codigo = "" & Trim(dxDBGrid1.Columns.ColumnByFieldName("idlote").Value)
    Descripcion = "" & Trim(dxDBGrid1.Columns.ColumnByFieldName("idlote").Value)
    DescripcionLote = "" & Trim(dxDBGrid1.Columns.ColumnByFieldName("GlsLote").Value)
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

Public Sub mostrar_from(indVale As String, Descripcion_lote As String, codigo_lote As String, codproducto As String, codalmacen As String)

    fill codproducto, codalmacen, indVale
    FrmAyudaLotes_Vales.Show 1
    codigo_lote = "" & codigo
    Descripcion_lote = "" & DescripcionLote
    Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub txtbusqueda_Change()
    
    If sw_limpia = False Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "idLote LIKE '*" & txtbusqueda.Text & "*' OR " & " GlsLote LIKE '*" & txtbusqueda.Text & "*'"
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
            dxDBGrid1.Dataset.Filter = "idLote LIKE '*" & txtbusqueda.Text & "*' OR " & " GlsLote LIKE '*" & txtbusqueda.Text & "*'"
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub
