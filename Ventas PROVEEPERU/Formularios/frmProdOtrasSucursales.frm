VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form frmProdOtrasSucursales 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos en otras Sucursales"
   ClientHeight    =   5220
   ClientLeft      =   1215
   ClientTop       =   1635
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBotones 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   30
      TabIndex        =   2
      Top             =   4410
      Width           =   11145
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1140
      End
   End
   Begin VB.Frame fraGrilla 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11145
      Begin DXDBGRIDLibCtl.dxDBGrid g 
         Height          =   4125
         Left            =   75
         OleObjectBlob   =   "frmProdOtrasSucursales.frx":0000
         TabIndex        =   1
         Top             =   150
         Width           =   10965
      End
   End
End
Attribute VB_Name = "frmProdOtrasSucursales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub MostrarForm(ByVal strCodProd As String, ByRef StrMsgError As String)
On Error GoTo Err
    
    MousePointer = 0
    ListaProductos strCodProd, StrMsgError
    If StrMsgError <> "" Then GoTo Err
    Me.Show vbModal
    Set G.DataSource = Nothing
    
    Unload Me
    
    Exit Sub
    
Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub ListaProductos(strCodProd As String, ByRef StrMsgError As String)
On Error GoTo Err
Dim strCampoUM As String
Dim strStockUM As String
Dim strSQL As String
Dim strCantidad As String
Dim strTablaPresentaciones As String

    strCampoUM = "idUMVenta"
    strStockUM = "CantidadStockUV"
    strCantidad = "(a.CantidadStock / f.Factor )" 'Es la cantidad de venta
    strTablaPresentaciones = " INNSER JOIN presentaciones f ON p.idEmpresa = f.idEmpresa AND p.idProducto = f.idProducto AND p." & strCampoUM & " = f.idUM "
    
    If indUMVenta = False Then
        strCampoUM = "idUMCompra"
        strStockUM = "CantidadStockUC"
        strCantidad = "a.CantidadStock" 'Es la cantidad de compra
        strTablaPresentaciones = ""
    End If
    
    strSQL = "SELECT concat(a.idSucursal,p.idProducto) AS Item,a.idSucursal,c.GlsPersona as GlsSucursal,p.idProducto,p.GlsProducto,m.GlsMarca,p." & strCampoUM & " AS idUMVenta,u.GlsUM,o.GlsMoneda,if(p.afectoIGV = 1,'S','N') Afecto, " & strCantidad & " as Stock, t.GlsTallaPeso " & _
             "FROM productos p " & _
                    "INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca " & _
                    "INNER JOIN unidadMedida u ON p." & strCampoUM & " = u.idUM " & _
                    "INNER JOIN monedas o ON p.idMoneda  = o.idMoneda " & _
                    "INNER JOIN productosalmacen a ON p.idEmpresa = a.idEmpresa AND a.idAlmacen = ( SELECT v.idAlmacen FROM almacenesvtas v WHERE v.idEmpresa = p.idEmpresa AND v.idSucursal = a.idSucursal ) AND p.idProducto = a.idProducto " & _
                    "INNER JOIN personas c ON a.idSucursal  = c.idPersona " & strTablaPresentaciones & _
                    "LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso " & _
             "WHERE  " & _
               " a.idSucursal <> '" & glsSucursal & "' " & _
               "AND p.idProducto = '" & strCodProd & "'"
                       
    With G
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strcn
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "Item"
    End With
    
    G.m.ClearGroupColumns
    G.Columns.ColumnByName("GlsSucursal").Caption = "Sucursal:"
    G.Columns.ColumnByName("GlsSucursal").GroupIndex = 0
    
    Exit Sub

Err:
    If StrMsgError = "" Then StrMsgError = Err.Description
End Sub

Private Sub Command1_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    ConfGrid G, False, False, False, False

End Sub

Private Sub g_OnReloadGroupList()
    
    G.m.FullExpand

End Sub
